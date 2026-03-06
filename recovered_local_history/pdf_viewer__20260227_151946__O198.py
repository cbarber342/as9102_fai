from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QGraphicsView, QGraphicsScene,
                               QGraphicsPixmapItem, QGraphicsItem, QGraphicsRectItem, QGraphicsLineItem,
                               QGraphicsTextItem, QPushButton, QInputDialog, QMenu, QMessageBox, QLabel,
                               QSlider, QComboBox, QDialog, QTextEdit, QFileDialog, QDialogButtonBox,
                               QSpinBox, QFormLayout, QApplication)
from PySide6.QtCore import Qt, QPointF, QRectF, Signal, QTimer, QSettings, QMimeData
from PySide6.QtGui import (QDragEnterEvent, QDropEvent, QPixmap, QImage, QBrush, QPen, 
                           QColor, QFont, QPainter, QPainterPath, QWheelEvent, QKeySequence, QMouseEvent, QShortcut, QTransform)
import fitz  # PyMuPDF
import re
import os
import json
import logging
from PIL import Image, ImageEnhance


logger = logging.getLogger(__name__)


class NotesExtractDialog(QDialog):
    insert_to_form3_requested = Signal(str, object)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Extracted Notes")
        self.resize(800, 600)

        layout = QVBoxLayout(self)

        # Action row (replaces right-click Insert-to-Form-3)
        top_row = QWidget(self)
        top_row_l = QHBoxLayout(top_row)
        top_row_l.setContentsMargins(0, 0, 0, 0)
        top_row_l.setSpacing(6)

        self.insert_selected_btn = QPushButton("Insert Selected to Form 3")
        self.insert_selected_btn.setEnabled(False)
        top_row_l.addWidget(self.insert_selected_btn)
        top_row_l.addStretch(1)

        self.source_label = QLabel("")
        try:
            self.source_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        except Exception:
            pass
        top_row_l.addWidget(self.source_label)
        layout.addWidget(top_row)

        self.text = QTextEdit()
        self.text.setReadOnly(False)

        # Keep default QTextEdit context menu (no custom "Insert to Form 3" item)

        layout.addWidget(self.text)

        def _update_enabled(_=None) -> None:
            try:
                self.insert_selected_btn.setEnabled(bool(self._selected_text()))
            except Exception:
                pass

        def _trigger_insert() -> None:
            s = self._selected_text()
            if s:
                try:
                    self.insert_to_form3_requested.emit(s, self)
                except Exception:
                    pass

        try:
            self.text.copyAvailable.connect(lambda _ok: _update_enabled())
        except Exception:
            pass
        try:
            self.text.cursorPositionChanged.connect(_update_enabled)
        except Exception:
            pass
        try:
            self.insert_selected_btn.clicked.connect(_trigger_insert)
        except Exception:
            pass

    def _selected_text(self) -> str:
        try:
            cur = self.text.textCursor()
            s = cur.selectedText() or ""
        except Exception:
            s = ""
        # QTextEdit uses U+2029 for newlines in selectedText().
        s = str(s).replace("\u2029", "\n").strip()
        return s

    def set_content(self, content: str) -> None:
        self.text.setPlainText(content or "")
        try:
            self.insert_selected_btn.setEnabled(bool(self._selected_text()))
        except Exception:
            pass

    def append_content(self, content: str) -> None:
        new_text = str(content or "").strip()
        if not new_text:
            return

        try:
            existing = self.text.toPlainText() or ""
        except Exception:
            existing = ""

        existing = str(existing)
        if existing.strip():
            combined = existing.rstrip() + "\n\n" + new_text
        else:
            combined = new_text
        self.text.setPlainText(combined)

        try:
            self.insert_selected_btn.setEnabled(bool(self._selected_text()))
        except Exception:
            pass

    def set_source(self, source: str) -> None:
        try:
            self.source_label.setText(str(source or ""))
        except Exception:
            pass

    def clear_content(self) -> None:
        try:
            self.text.clear()
        except Exception:
            pass
        try:
            self.insert_selected_btn.setEnabled(False)
        except Exception:
            pass
        try:
            self.source_label.setText("")
        except Exception:
            pass


class _NoteRegionItem(QGraphicsRectItem):
    """Resizable/movable note-region rectangle with visible drag handles."""

    HANDLE_SIZE = 8.0
    MIN_SIZE = 12.0

    def __init__(self, rect: QRectF, *, viewer: "PdfViewer", index0: int):
        super().__init__(rect)
        self._viewer = viewer
        self._index0 = int(index0)
        self._active_handle: str | None = None
        self._press_pos: QPointF | None = None
        self._press_rect: QRectF | None = None
        self._hover = False

        self.setAcceptHoverEvents(True)
        self.setZValue(60)
        self.setFlag(QGraphicsItem.ItemIsSelectable, True)
        self.setFlag(QGraphicsItem.ItemIsMovable, True)
        self.setFlag(QGraphicsItem.ItemSendsGeometryChanges, True)

    def _handle_rects(self) -> dict[str, QRectF]:
        r = self.rect()
        s = float(self.HANDLE_SIZE)
        hs = s / 2.0
        cx = r.center().x()
        cy = r.center().y()
        return {
            "tl": QRectF(r.left() - hs, r.top() - hs, s, s),
            "tm": QRectF(cx - hs, r.top() - hs, s, s),
            "tr": QRectF(r.right() - hs, r.top() - hs, s, s),
            "ml": QRectF(r.left() - hs, cy - hs, s, s),
            "mr": QRectF(r.right() - hs, cy - hs, s, s),
            "bl": QRectF(r.left() - hs, r.bottom() - hs, s, s),
            "bm": QRectF(cx - hs, r.bottom() - hs, s, s),
            "br": QRectF(r.right() - hs, r.bottom() - hs, s, s),
        }

    def boundingRect(self) -> QRectF:
        # Expand bounds so handles are clickable (handles extend outside the rect).
        try:
            hs = float(self.HANDLE_SIZE) / 2.0
        except Exception:
            hs = 4.0
        return super().boundingRect().adjusted(-hs, -hs, hs, hs)

    def shape(self):
        # Expand hit-test area to include handles.
        p = QPainterPath()
        p.addRect(self.boundingRect())
        return p

    def _hit_handle(self, pos: QPointF) -> str | None:
        # First: direct handle hit.
        try:
            for name, hr in self._handle_rects().items():
                if hr.contains(pos):
                    return name
        except Exception:
            pass

        # Second: allow grabbing edges/mid-sides by proximity.
        try:
            r = self.rect()
            thr = float(self.HANDLE_SIZE)
            near_l = abs(pos.x() - r.left()) <= thr
            near_r = abs(pos.x() - r.right()) <= thr
            near_t = abs(pos.y() - r.top()) <= thr
            near_b = abs(pos.y() - r.bottom()) <= thr

            if near_l and near_t:
                return "tl"
            if near_r and near_t:
                return "tr"
            if near_l and near_b:
                return "bl"
            if near_r and near_b:
                return "br"
            if near_t:
                return "tm"
            if near_b:
                return "bm"
            if near_l:
                return "ml"
            if near_r:
                return "mr"
        except Exception:
            pass
        return None

    def hoverMoveEvent(self, event):
        try:
            self._hover = True
            h = self._hit_handle(event.pos())
            if h in ("tl", "br"):
                self.setCursor(Qt.SizeFDiagCursor)
            elif h in ("tr", "bl"):
                self.setCursor(Qt.SizeBDiagCursor)
            elif h in ("tm", "bm"):
                self.setCursor(Qt.SizeVerCursor)
            elif h in ("ml", "mr"):
                self.setCursor(Qt.SizeHorCursor)
            else:
                self.setCursor(Qt.SizeAllCursor)
        except Exception:
            pass
        try:
            self.update()
        except Exception:
            pass
        return super().hoverMoveEvent(event)

    def hoverLeaveEvent(self, event):
        try:
            self._hover = False
            self.unsetCursor()
        except Exception:
            pass
        try:
            self.update()
        except Exception:
            pass
        return super().hoverLeaveEvent(event)

    def mousePressEvent(self, event):
        try:
            if event.button() == Qt.LeftButton:
                h = self._hit_handle(event.pos())
                if h is not None:
                    self._active_handle = h
                    self._press_pos = event.pos()
                    self._press_rect = QRectF(self.rect())
                    event.accept()
                    return
        except Exception:
            pass
        return super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._active_handle is None:
            return super().mouseMoveEvent(event)

        try:
            if self._press_pos is None or self._press_rect is None:
                return super().mouseMoveEvent(event)

            delta = event.pos() - self._press_pos
            r = QRectF(self._press_rect)

            left = r.left()
            right = r.right()
            top = r.top()
            bottom = r.bottom()

            if "l" in self._active_handle:
                left = left + delta.x()
            if "r" in self._active_handle:
                right = right + delta.x()
            if "t" in self._active_handle:
                top = top + delta.y()
            if "b" in self._active_handle:
                bottom = bottom + delta.y()

            # Enforce minimum size.
            if (right - left) < float(self.MIN_SIZE):
                if "l" in self._active_handle:
                    left = right - float(self.MIN_SIZE)
                else:
                    right = left + float(self.MIN_SIZE)
            if (bottom - top) < float(self.MIN_SIZE):
                if "t" in self._active_handle:
                    top = bottom - float(self.MIN_SIZE)
                else:
                    bottom = top + float(self.MIN_SIZE)

            # Clamp within the page bounds.
            try:
                page_rect = self._viewer.pixmap_item.boundingRect() if self._viewer.pixmap_item is not None else None
                if page_rect is not None:
                    left = max(page_rect.left(), min(left, page_rect.right() - float(self.MIN_SIZE)))
                    top = max(page_rect.top(), min(top, page_rect.bottom() - float(self.MIN_SIZE)))
                    right = max(page_rect.left() + float(self.MIN_SIZE), min(right, page_rect.right()))
                    bottom = max(page_rect.top() + float(self.MIN_SIZE), min(bottom, page_rect.bottom()))
            except Exception:
                pass

            self.setRect(QRectF(left, top, right - left, bottom - top))
            event.accept()
            return
        except Exception:
            return super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        try:
            if self._active_handle is not None:
                self._active_handle = None
                self._press_pos = None
                self._press_rect = None
        except Exception:
            pass
        try:
            # Sync the (possibly moved/resized) rect back to normalized storage.
            if self._viewer is not None:
                self._viewer._update_note_region_from_item(self._index0, self.mapRectToScene(self.rect()))
        except Exception:
            pass
        return super().mouseReleaseEvent(event)

    def paint(self, painter, option, widget=None):
        super().paint(painter, option, widget)
        try:
            if not (self.isSelected() or self._hover):
                return
            painter.save()
            painter.setPen(QPen(QColor(255, 200, 0), 1))
            painter.setBrush(QBrush(QColor(255, 200, 0)))
            for hr in self._handle_rects().values():
                painter.drawRect(hr)
            painter.restore()
        except Exception:
            return


class _GridBoundsItem(QGraphicsRectItem):
    """Interactive grid bounds rectangle with resize handles.

    This is only shown when the drawing grid is enabled. Drag corners/edges to
    resize, or drag inside to move. During drag it updates the overlay and
    emits viewer.grid_bounds_changed with percent bounds.
    """

    HANDLE_SIZE = 8.0
    MIN_SIZE = 24.0

    def __init__(self, rect: QRectF, *, viewer: "PdfViewer"):
        super().__init__(rect)
        self._viewer = viewer
        self._active_handle: str | None = None
        self._press_scene: QPointF | None = None
        self._press_rect: QRectF | None = None
        self._hover = False

        self.setAcceptHoverEvents(True)
        self.setZValue(52)
        self.setFlag(QGraphicsItem.ItemIsSelectable, True)
        self.setFlag(QGraphicsItem.ItemIsMovable, False)
        self.setFlag(QGraphicsItem.ItemSendsGeometryChanges, True)

        try:
            self.setCursor(Qt.SizeAllCursor)
        except Exception:
            pass

    def _handle_rects(self) -> dict[str, QRectF]:
        r = self.rect()
        s = float(self.HANDLE_SIZE)
        hs = s / 2.0
        cx = r.center().x()
        cy = r.center().y()
        return {
            "tl": QRectF(r.left() - hs, r.top() - hs, s, s),
            "tm": QRectF(cx - hs, r.top() - hs, s, s),
            "tr": QRectF(r.right() - hs, r.top() - hs, s, s),
            "ml": QRectF(r.left() - hs, cy - hs, s, s),
            "mr": QRectF(r.right() - hs, cy - hs, s, s),
            "bl": QRectF(r.left() - hs, r.bottom() - hs, s, s),
            "bm": QRectF(cx - hs, r.bottom() - hs, s, s),
            "br": QRectF(r.right() - hs, r.bottom() - hs, s, s),
        }

    def boundingRect(self) -> QRectF:
        try:
            hs = float(self.HANDLE_SIZE) / 2.0
        except Exception:
            hs = 4.0
        return super().boundingRect().adjusted(-hs, -hs, hs, hs)

    def shape(self):
        p = QPainterPath()
        p.addRect(self.boundingRect())
        return p

    def _hit_handle(self, pos: QPointF) -> str | None:
        try:
            for name, hr in self._handle_rects().items():
                if hr.contains(pos):
                    return name
        except Exception:
            pass

        # Proximity to edges
        try:
            r = self.rect()
            thr = float(self.HANDLE_SIZE)
            near_l = abs(pos.x() - r.left()) <= thr
            near_r = abs(pos.x() - r.right()) <= thr
            near_t = abs(pos.y() - r.top()) <= thr
            near_b = abs(pos.y() - r.bottom()) <= thr

            if near_l and near_t:
                return "tl"
            if near_r and near_t:
                return "tr"
            if near_l and near_b:
                return "bl"
            if near_r and near_b:
                return "br"
            if near_t:
                return "tm"
            if near_b:
                return "bm"
            if near_l:
                return "ml"
            if near_r:
                return "mr"
        except Exception:
            pass
        return None

    def hoverMoveEvent(self, event):
        try:
            self._hover = True
            h = self._hit_handle(event.pos())
            if h in ("tl", "br"):
                self.setCursor(Qt.SizeFDiagCursor)
            elif h in ("tr", "bl"):
                self.setCursor(Qt.SizeBDiagCursor)
            elif h in ("tm", "bm"):
                self.setCursor(Qt.SizeVerCursor)
            elif h in ("ml", "mr"):
                self.setCursor(Qt.SizeHorCursor)
            else:
                self.setCursor(Qt.SizeAllCursor)
        except Exception:
            pass
        try:
            self.update()
        except Exception:
            pass
        return super().hoverMoveEvent(event)

    def hoverLeaveEvent(self, event):
        try:
            self._hover = False
            self.unsetCursor()
        except Exception:
            pass
        try:
            self.update()
        except Exception:
            pass
        return super().hoverLeaveEvent(event)

    def mousePressEvent(self, event):
        try:
            if event.button() != Qt.LeftButton:
                return super().mousePressEvent(event)
            self._active_handle = self._hit_handle(event.pos())
            self._press_scene = event.scenePos()
            self._press_rect = QRectF(self.rect())
            event.accept()
            return
        except Exception:
            return super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        try:
            if self._press_scene is None or self._press_rect is None:
                return super().mouseMoveEvent(event)

            page_rect = None
            try:
                page_rect = self._viewer.pixmap_item.boundingRect() if self._viewer.pixmap_item is not None else None
            except Exception:
                page_rect = None
            if page_rect is None:
                return super().mouseMoveEvent(event)

            d = event.scenePos() - self._press_scene
            r0 = QRectF(self._press_rect)

            left = float(r0.left())
            top = float(r0.top())
            right = float(r0.right())
            bottom = float(r0.bottom())

            h = self._active_handle
            if h is None:
                # Move
                left += float(d.x())
                right += float(d.x())
                top += float(d.y())
                bottom += float(d.y())
            else:
                if "l" in h:
                    left += float(d.x())
                if "r" in h:
                    right += float(d.x())
                if "t" in h:
                    top += float(d.y())
                if "b" in h:
                    bottom += float(d.y())

            # Enforce minimum size
            if (right - left) < float(self.MIN_SIZE):
                if h and "l" in h:
                    left = right - float(self.MIN_SIZE)
                else:
                    right = left + float(self.MIN_SIZE)
            if (bottom - top) < float(self.MIN_SIZE):
                if h and "t" in h:
                    top = bottom - float(self.MIN_SIZE)
                else:
                    bottom = top + float(self.MIN_SIZE)

            # Clamp within page
            if h is None:
                w = right - left
                hgt = bottom - top
                left = max(page_rect.left(), min(left, page_rect.right() - w))
                top = max(page_rect.top(), min(top, page_rect.bottom() - hgt))
                right = left + w
                bottom = top + hgt
            else:
                left = max(page_rect.left(), min(left, page_rect.right() - float(self.MIN_SIZE)))
                top = max(page_rect.top(), min(top, page_rect.bottom() - float(self.MIN_SIZE)))
                right = max(page_rect.left() + float(self.MIN_SIZE), min(right, page_rect.right()))
                bottom = max(page_rect.top() + float(self.MIN_SIZE), min(bottom, page_rect.bottom()))

            new_rect = QRectF(left, top, right - left, bottom - top)
            self.setRect(new_rect)
            try:
                self._viewer._on_grid_bounds_rect_dragging(QRectF(new_rect))
            except Exception:
                pass
            event.accept()
            return
        except Exception:
            return super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        try:
            if self._press_rect is not None:
                try:
                    self._viewer._on_grid_bounds_rect_committed(QRectF(self.rect()))
                except Exception:
                    pass
        except Exception:
            pass
        try:
            self._active_handle = None
            self._press_scene = None
            self._press_rect = None
        except Exception:
            pass
        return super().mouseReleaseEvent(event)

    def paint(self, painter, option, widget=None):
        # Outline
        try:
            painter.save()
            painter.setBrush(Qt.NoBrush)
            painter.setPen(QPen(QColor(0, 0, 0, 200), 1, Qt.DashLine))
            painter.drawRect(self.rect())

            # Handles only when selected/hovered
            if self.isSelected() or self._hover:
                painter.setPen(QPen(QColor(0, 0, 0, 220), 1))
                painter.setBrush(QBrush(QColor(255, 255, 255, 220)))
                for hr in self._handle_rects().values():
                    painter.drawRect(hr)
            painter.restore()
        except Exception:
            pass



class BubbleItem(QGraphicsItem):
    """A draggable bubble annotation with a number inside."""
    
    def __init__(self, number, x, y, base_radius=15, parent_viewer=None, range_end: int | None = None, display_text: str | None = None, backfill_rgb: str | None = None):
        super().__init__()
        self.number = number
        self.range_end = int(range_end) if range_end is not None else int(number)
        self.text = str(display_text) if display_text is not None else self._format_text()
        self.setPos(x, y)
        self.base_radius = base_radius
        self.parent_viewer = parent_viewer
        self.backfill_rgb = str(backfill_rgb).strip().upper() if backfill_rgb else ""

        self.setFlags(QGraphicsItem.ItemIsMovable | QGraphicsItem.ItemIsSelectable)
        self.setAcceptedMouseButtons(Qt.LeftButton)
        self.setAcceptHoverEvents(True)
        self.setZValue(100)  # Keep bubbles on top
        self.setCursor(Qt.OpenHandCursor)
        self._drag_offset_scene: QPointF | None = None
        self._group_drag_start_scene: QPointF | None = None
        self._group_drag_positions: dict["BubbleItem", QPointF] | None = None

    def _format_text(self) -> str:
        try:
            start = int(self.number)
            end = int(self.range_end)
        except Exception:
            return str(self.number)
        if end > start:
            return f"{start}-{end}"
        return str(start)

    def _width_multiplier_for_text(self) -> float:
        # Expand width only (not height) so long labels fit.
        text = str(getattr(self, "text", "") or "")
        n = len(text)
        if n <= 2:
            return 1.0
        if n == 3:
            return 1.20
        if n == 4:
            return 1.35
        if n == 5:
            return 1.55
        if n == 6:
            return 1.75
        if n == 7:
            return 1.95
        if n <= 9:
            return 2.50
        return 3.00

    def _half_sizes(self) -> tuple[float, float]:
        # Returns (half_width, half_height) in scene units.
        zoom = 1.0
        if self.parent_viewer is not None:
            try:
                zoom = float(self.parent_viewer.get_zoom_factor())
            except Exception:
                zoom = 1.0
        half_h = float(self.base_radius) * zoom
        half_w = float(self.base_radius) * self._width_multiplier_for_text() * zoom
        return (half_w, half_h)

    def _effective_line_width_f(self) -> float:
        """Return the pen width (in scene units) used for on-screen rendering.

        Zoom in this viewer is implemented by re-rendering the PDF pixmap (scene gets larger/smaller)
        rather than scaling the QGraphicsView transform. Bubble geometry scales with zoom, but the
        outline width must also scale with zoom to avoid becoming disproportionately thick when
        zooming out (where text can appear to disappear).
        """
        base_lw = 3.0
        zoom = 1.0
        if self.parent_viewer is not None:
            try:
                base_lw = float(getattr(self.parent_viewer, "bubble_line_width", base_lw) or base_lw)
            except Exception:
                base_lw = 3.0
            try:
                zoom = float(self.parent_viewer.get_zoom_factor())
            except Exception:
                zoom = 1.0
        # Keep visible at low zoom but allow thinning.
        return max(0.5, base_lw * max(0.25, zoom))
        
    @property
    def radius(self):
        # Keep legacy meaning: half-height (controls feel unchanged).
        if self.parent_viewer is not None:
            return self.base_radius * self.parent_viewer.get_zoom_factor()
        return self.base_radius
        
    def boundingRect(self):
        half_w, half_h = self._half_sizes()
        pad = 3.0
        # Include stroke width so thick outlines are not clipped.
        line_width_f = self._effective_line_width_f()
        stroke_pad = max(pad, float(line_width_f) / 2.0 + 2.0)
        return QRectF(
            -(half_w + stroke_pad),
            -(half_h + stroke_pad),
            (half_w + stroke_pad) * 2,
            (half_h + stroke_pad) * 2,
        )
        
    def paint(self, painter, option, widget):
        painter.setRenderHint(QPainter.Antialiasing, True)

        half_w, half_h = self._half_sizes()
        bubble_shape = "Circle"
        line_width_f = self._effective_line_width_f()
        if self.parent_viewer is not None:
            bubble_shape = getattr(self.parent_viewer, "bubble_shape", bubble_shape)
        
        # Draw bubble outline - configurable outline color, no fill (transparent)
        base_color = QColor(220, 40, 40)
        if self.parent_viewer is not None:
            try:
                c = getattr(self.parent_viewer, "bubble_color", None)
                if c is not None and c.isValid():
                    base_color = QColor(c)
            except Exception:
                pass
        bubble_fill = None
        try:
            bf = str(getattr(self, "backfill_rgb", "") or "").strip()
        except Exception:
            bf = ""
        if bf:
            try:
                qc = QColor("#" + bf)
                if qc.isValid():
                    bubble_fill = qc
            except Exception:
                bubble_fill = None
        if bubble_fill is None and self.parent_viewer is not None:
            try:
                if hasattr(self.parent_viewer, "get_bubble_backfill_qcolor"):
                    bubble_fill = self.parent_viewer.get_bubble_backfill_qcolor()
                else:
                    bubble_fill_white = bool(getattr(self.parent_viewer, "bubble_backfill_white", False))
                    bubble_fill = QColor("#FFFFFF") if bubble_fill_white else None
            except Exception:
                bubble_fill = None

        if self.isSelected():
            painter.setBrush(QBrush(QColor(0, 120, 255, 50)))
            pen = QPen(QColor(0, 120, 255))
            pen.setWidthF(max(0.5, line_width_f + 1.0))
            painter.setPen(pen)
        else:
            painter.setBrush(QBrush(bubble_fill) if bubble_fill is not None else QBrush(Qt.transparent))
            pen = QPen(base_color)
            pen.setWidthF(max(0.5, line_width_f))
            painter.setPen(pen)

        if str(bubble_shape).lower().startswith("rect"):
            # Draw the outline slightly *outside* the nominal bubble bounds so thick
            # strokes don't consume the interior and hide the text.
            outline_half_w = float(half_w) + float(line_width_f) / 2.0
            outline_half_h = float(half_h) + float(line_width_f) / 2.0
            painter.drawRect(QRectF(-outline_half_w, -outline_half_h, outline_half_w * 2.0, outline_half_h * 2.0))
        else:
            outline_half_w = float(half_w) + float(line_width_f) / 2.0
            outline_half_h = float(half_h) + float(line_width_f) / 2.0
            painter.drawEllipse(QRectF(-outline_half_w, -outline_half_h, outline_half_w * 2.0, outline_half_h * 2.0))
        
        # Draw text - match outline color
        if self.isSelected():
            painter.setPen(QPen(QColor(0, 120, 255)))
        else:
            painter.setPen(QPen(base_color))
        text = self.text
        
        # Calculate font size to fit inside bubble.
        # Since the outline is drawn outside, only keep a small margin.
        inner_w = max(4.0, float(half_w) - 0.5)
        inner_h = max(4.0, float(half_h) - 0.5)

        start_font = max(5, int(inner_h * 1.4))
        # If no font fits, keep shrinking down to a small minimum.
        min_font = 3
        font_size = min_font
        for fs in range(start_font, min_font - 1, -1):
            font = QFont("Arial", fs, QFont.Bold)
            painter.setFont(font)
            fm = painter.fontMetrics()
            text_width = fm.horizontalAdvance(text)
            text_height = fm.height()
            if text_width < inner_w * 1.8 and text_height < inner_h * 1.5:
                font_size = fs
                break
        painter.setFont(QFont("Arial", int(font_size), QFont.Bold))

        # Draw centered text
        painter.drawText(QRectF(-inner_w, -inner_h, inner_w * 2, inner_h * 2), Qt.AlignCenter, text)

    def set_base_radius(self, new_base_radius):
        self.base_radius = new_base_radius
        self.prepareGeometryChange()
        self.update()

    def set_number(self, number):
        self.prepareGeometryChange()
        self.number = number
        self.range_end = int(number)
        self.text = self._format_text()
        self.update()

    def set_range_end(self, end: int) -> None:
        self.prepareGeometryChange()
        self.range_end = int(end)
        self.text = self._format_text()
        self.update()

    def contextMenuEvent(self, event):
        menu = QMenu()
        resize_action = menu.addAction("Resize...")
        renumber_action = menu.addAction("Renumber...")
        delete_action = menu.addAction("Delete")
        
        action = menu.exec(event.screenPos())
        
        if action == delete_action and self.parent_viewer:
            self.parent_viewer.remove_bubble(self)
        elif action == renumber_action and self.parent_viewer:
            self.parent_viewer.renumber_bubble(self)
        elif action == resize_action and self.parent_viewer:
            self.parent_viewer.resize_bubble(self)

    def mousePressEvent(self, event):
        if self.parent_viewer is not None:
            try:
                self.parent_viewer._debug(
                    f"BubbleItem.mousePressEvent: button={int(event.button())} start={self.number} end={getattr(self, 'range_end', None)}"
                )
            except Exception:
                pass
        suppress_undo = bool(getattr(self, "_suppress_next_press_undo", False))
        if suppress_undo:
            try:
                self._suppress_next_press_undo = False
            except Exception:
                pass
        if (not suppress_undo) and self.parent_viewer is not None:
            try:
                self.parent_viewer._push_undo_state()
            except Exception:
                pass
        if event.button() == Qt.LeftButton:
            try:
                # If multiple bubbles are selected, drag them as a group.
                scene = self.scene()
                selected = []
                if scene is not None:
                    try:
                        selected = [it for it in (scene.selectedItems() or []) if isinstance(it, BubbleItem)]
                    except Exception:
                        selected = []
                if len(selected) > 1 and self in selected:
                    self._group_drag_start_scene = event.scenePos()
                    try:
                        self._group_drag_positions = {it: QPointF(it.pos()) for it in selected}
                    except Exception:
                        self._group_drag_positions = None
                    self._drag_offset_scene = None
                else:
                    # Store offset in scene coordinates for robust manual dragging.
                    self._drag_offset_scene = event.scenePos() - self.scenePos()
            except Exception:
                self._drag_offset_scene = None
        
        # Track start position to detect if a move actually occurred
        try:
            self._drag_start_pos = self.scenePos()
        except Exception:
            self._drag_start_pos = None

        try:
            self.setCursor(Qt.ClosedHandCursor)
        except Exception:
            pass
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.parent_viewer is not None:
            try:
                self.parent_viewer._debug(
                    f"BubbleItem.mouseMoveEvent: drag_offset={'set' if self._drag_offset_scene is not None else 'none'} scenePos=({event.scenePos().x():.1f},{event.scenePos().y():.1f})"
                )
            except Exception:
                pass
        # Group-drag selected bubbles together.
        if self._group_drag_start_scene is not None and self._group_drag_positions is not None:
            try:
                delta = event.scenePos() - self._group_drag_start_scene
                for it, p0 in self._group_drag_positions.items():
                    try:
                        it.setPos(p0 + delta)
                    except Exception:
                        continue
                event.accept()
                return
            except Exception:
                pass

        # Manual drag for reliability.
        if self._drag_offset_scene is not None:
            try:
                new_pos = event.scenePos() - self._drag_offset_scene
                self.setPos(new_pos)
                event.accept()
                return
            except Exception:
                pass
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self.parent_viewer is not None:
            try:
                self.parent_viewer._debug(
                    f"BubbleItem.mouseReleaseEvent: button={int(event.button())} pos=({self.pos().x():.1f},{self.pos().y():.1f})"
                )
            except Exception:
                pass
        super().mouseReleaseEvent(event)
        self._drag_offset_scene = None
        self._group_drag_start_scene = None
        self._group_drag_positions = None
        try:
            self.setCursor(Qt.OpenHandCursor)
        except Exception:
            pass
        if self.parent_viewer is not None:
            try:
                # Auto-resolve overlaps after manual moves, but ONLY if the item actually moved.
                moved_significantly = True
                try:
                    if hasattr(self, "_drag_start_pos") and self._drag_start_pos is not None:
                        diff = self.scenePos() - self._drag_start_pos
                        if diff.manhattanLength() < 0.5:
                            moved_significantly = False
                except Exception:
                    pass

                try:
                    if moved_significantly and hasattr(self.parent_viewer, "auto_resolve_bubble_overlaps_current_page"):
                        self.parent_viewer.auto_resolve_bubble_overlaps_current_page()
                except Exception:
                    pass
                self.parent_viewer._persist_current_page_bubbles()
                # Force update of Form 3 zones when a bubble is dropped/moved
                try:
                    self.parent_viewer.bubbles_changed.emit(self.parent_viewer.get_bubbled_numbers())
                except Exception:
                    pass
            except Exception:
                pass

    def hoverEnterEvent(self, event):
        try:
            self.setCursor(Qt.OpenHandCursor)
        except Exception:
            pass
        super().hoverEnterEvent(event)

    def hoverLeaveEvent(self, event):
        try:
            self.setCursor(Qt.OpenHandCursor)
        except Exception:
            pass
        super().hoverLeaveEvent(event)


class InteractiveGraphicsView(QGraphicsView):
    """Custom graphics view with zoom, pan, and bubble placement."""
    
    bubble_click = Signal(QPointF)
    note_region_created = Signal(QRectF)
    
    def __init__(self, scene, parent=None):
        super().__init__(scene, parent)
        self.parent_viewer = parent
        self.placing_bubble = False
        self.placing_note_region = False
        self._note_drag_start = None
        self._note_drag_item = None
        
        # High quality rendering
        self.setRenderHint(QPainter.Antialiasing, True)
        self.setRenderHint(QPainter.SmoothPixmapTransform, True)
        self.setRenderHint(QPainter.TextAntialiasing, True)
        
        # Default: rubber-band selection for "select window".
        # Item dragging still works when the press begins on a bubble.
        # We'll pan with middle mouse.
        self.setDragMode(QGraphicsView.RubberBandDrag)
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.AnchorUnderMouse)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        
        self.setAcceptDrops(True)
        self.setMouseTracking(True)
        self.setBackgroundBrush(QBrush(QColor(70, 70, 70)))
        self.setInteractive(True)

        self._panning = False
        self._pan_last_pos = None

    def set_placing_mode(self, placing):
        """Enable/disable bubble placement mode."""
        self.placing_bubble = placing
        if placing:
            self.setCursor(Qt.CrossCursor)
            self.setDragMode(QGraphicsView.NoDrag)
        else:
            self.setCursor(Qt.ArrowCursor)
            self.setDragMode(QGraphicsView.RubberBandDrag)

    def set_note_region_mode(self, placing: bool) -> None:
        """Enable/disable note region drag-to-select mode."""
        self.placing_note_region = placing
        if placing:
            self.setCursor(Qt.CrossCursor)
            self.setDragMode(QGraphicsView.NoDrag)
        else:
            self.setCursor(Qt.ArrowCursor)
            self.setDragMode(QGraphicsView.RubberBandDrag)

    def mousePressEvent(self, event: QMouseEvent):
        if self.parent_viewer is not None:
            try:
                self.parent_viewer._debug(
                    f"View.mousePressEvent: button={int(event.button())} placing_bubble={self.placing_bubble} placing_note={self.placing_note_region}"
                )
            except Exception:
                pass

        # If the click is on an existing bubble or Notes Window, always let the item handle it
        # (drag/select/resize) even if we're currently in placement mode.
        if event.button() == Qt.LeftButton:
            try:
                item = self.itemAt(event.pos())
            except Exception:
                item = None

            # Extra hit-test visibility when debugging
            if self.parent_viewer is not None and getattr(self.parent_viewer, "_debug_enabled", False):
                try:
                    items_here = self.items(event.pos())
                    names = []
                    for it in items_here[:5]:
                        try:
                            names.append(type(it).__name__)
                        except Exception:
                            names.append("?")
                    self.parent_viewer._debug(f"View.itemsAt: {names}")
                except Exception:
                    pass
            bubble_item = item
            try:
                while bubble_item is not None and not isinstance(bubble_item, BubbleItem):
                    bubble_item = bubble_item.parentItem()
            except Exception:
                bubble_item = None
            if bubble_item is not None:
                if self.parent_viewer is not None:
                    try:
                        self.parent_viewer._debug("View: click over bubble; passing through to item")
                    except Exception:
                        pass
                return super().mousePressEvent(event)

            # Notes Window item pass-through
            note_item = item
            try:
                while note_item is not None and not isinstance(note_item, _NoteRegionItem):
                    note_item = note_item.parentItem()
            except Exception:
                note_item = None
            if note_item is not None:
                if self.parent_viewer is not None:
                    try:
                        self.parent_viewer._debug("View: click over note region; passing through to item")
                    except Exception:
                        pass
                return super().mousePressEvent(event)

        # Middle mouse pans the view without interfering with item dragging.
        if event.button() == Qt.MiddleButton:
            self._panning = True
            self._pan_last_pos = event.pos()
            self.setCursor(Qt.ClosedHandCursor)
            event.accept()
            return
        if self.placing_note_region and event.button() == Qt.LeftButton:
            self._note_drag_start = self.mapToScene(event.pos())
            if self._note_drag_item is not None:
                try:
                    self.scene().removeItem(self._note_drag_item)
                except Exception:
                    pass
            self._note_drag_item = QGraphicsRectItem(QRectF(self._note_drag_start, self._note_drag_start))
            self._note_drag_item.setPen(QPen(QColor(255, 200, 0), 2))
            self._note_drag_item.setBrush(QBrush(QColor(255, 200, 0, 40)))
            self._note_drag_item.setZValue(50)
            self.scene().addItem(self._note_drag_item)
            event.accept()
            return
        if self.placing_bubble and event.button() == Qt.LeftButton:
            scene_pos = self.mapToScene(event.pos())
            if self.parent_viewer is not None:
                try:
                    self.parent_viewer._debug(
                        f"View: emitting bubble_click at ({scene_pos.x():.1f},{scene_pos.y():.1f})"
                    )
                except Exception:
                    pass
            self.bubble_click.emit(scene_pos)
            # Important: do NOT swallow the press.
            # Forward the same press to the scene so the newly created BubbleItem
            # can immediately receive it and be dragged.
            return super().mousePressEvent(event)
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event: QMouseEvent):
        if self._panning and self._pan_last_pos is not None:
            delta = event.pos() - self._pan_last_pos
            self._pan_last_pos = event.pos()
            self.horizontalScrollBar().setValue(self.horizontalScrollBar().value() - delta.x())
            self.verticalScrollBar().setValue(self.verticalScrollBar().value() - delta.y())
            event.accept()
            return
        if self.placing_note_region and self._note_drag_start is not None and self._note_drag_item is not None:
            current = self.mapToScene(event.pos())
            rect = QRectF(self._note_drag_start, current).normalized()
            self._note_drag_item.setRect(rect)
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event: QMouseEvent):
        if event.button() == Qt.MiddleButton and self._panning:
            self._panning = False
            self._pan_last_pos = None
            # Restore cursor based on current mode
            if self.placing_bubble or self.placing_note_region:
                self.setCursor(Qt.CrossCursor)
            else:
                self.setCursor(Qt.ArrowCursor)
            event.accept()
            return
        if self.placing_note_region and event.button() == Qt.LeftButton and self._note_drag_start is not None:
            if self._note_drag_item is not None:
                rect = self._note_drag_item.rect().normalized()
                try:
                    self.scene().removeItem(self._note_drag_item)
                except Exception:
                    pass
                self._note_drag_item = None
                self._note_drag_start = None

                # Ignore tiny drags
                if rect.width() >= 10 and rect.height() >= 10:
                    self.note_region_created.emit(rect)

            event.accept()
            return
        super().mouseReleaseEvent(event)

    def wheelEvent(self, event: QWheelEvent):
        # Zoom with Ctrl + scroll wheel (re-rendered for crispness)
        if self.parent_viewer and (event.modifiers() & Qt.ControlModifier):
            try:
                delta_y = float(event.angleDelta().y())
            except Exception:
                delta_y = 0.0

            # Typical mouse wheel notch is 120. Use an exponential scale so zoom speed
            # feels responsive, and also works well with smooth trackpads.
            try:
                steps = delta_y / 120.0
            except Exception:
                steps = 0.0

            # Larger base factor = faster zoom per wheel notch.
            base = 1.25
            try:
                factor = pow(base, steps)
            except Exception:
                factor = base if steps > 0 else (1.0 / base)

            try:
                cur = float(self.parent_viewer.get_zoom_factor())
            except Exception:
                cur = 1.0
            self.parent_viewer.set_zoom(cur * factor)
            event.accept()
            return
        # Page scroll with wheel (when not zooming): if at edge, move page.
        if self.parent_viewer is not None:
            try:
                delta_y = int(event.angleDelta().y())
            except Exception:
                delta_y = 0

            try:
                vbar = self.verticalScrollBar()
                vmin = int(vbar.minimum())
                vmax = int(vbar.maximum())
                vval = int(vbar.value())
            except Exception:
                vbar = None
                vmin = vmax = vval = 0

            at_top = (vval <= vmin + 1)
            at_bottom = (vval >= vmax - 1)
            no_scroll = (vmax <= vmin)

            try:
                if delta_y < 0 and (at_bottom or no_scroll):
                    self.parent_viewer.next_page()
                    event.accept()
                    return
                if delta_y > 0 and (at_top or no_scroll):
                    self.parent_viewer.prev_page()
                    event.accept()
                    return
            except Exception:
                pass
        super().wheelEvent(event)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls() or event.mimeData().hasText():
            event.accept()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        event.accept()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            self.parent_viewer.handle_file_drop(event)
        elif event.mimeData().hasText():
            text = event.mimeData().text()
            scene_pos = self.mapToScene(event.position().toPoint())
            self.parent_viewer.add_bubble_from_drop(text, scene_pos.x(), scene_pos.y())
            event.accept()


class PdfViewer(QWidget):
    """PDF viewer with multi-page support and bubble annotations."""
    
    bubble_added = Signal(int)
    bubble_removed = Signal(int)
    # Emitted whenever the set of bubble numbers present changes.
    # Payload is a `set[int]`.
    bubbles_changed = Signal(object)
    # Emitted after a successful save; payload is the output PDF path.
    drawing_saved = Signal(str)

    # Emitted when interactive grid bounds are dragged/resized.
    # Payload: (left_pct, top_pct, width_pct, height_pct)
    grid_bounds_changed = Signal(float, float, float, float)

    # Emitted when user requests inserting selected extracted notes into Form 3 column G.
    insert_notes_to_form3_requested = Signal(str, object)
    
    def __init__(
        self,
        high_res: bool = False,
        default_save_basename: str | None = None,
        *,
        embed_controls: bool = True,
    ):
        super().__init__()
        # Fullscreen/high-res mode removed per UX request. Keep arg for compatibility.
        self.high_res = False
        self.default_save_basename = default_save_basename or ""
        self._embed_controls = bool(embed_controls)
        self._debug_enabled = str(os.environ.get("AS9102_DEBUG_PDF", "")).strip().lower() in ("1", "true", "yes", "on")
        self.setFocusPolicy(Qt.StrongFocus)

        # Persist viewer preferences across runs.
        self._settings = QSettings("as9102_fai", "as9102_fai_gui")
        self._size_slider_value = int(self._settings.value("pdf_viewer/bubble_size", 5, int) or 5)
        self._size_slider_value = max(1, min(10, self._size_slider_value))
        self._line_slider_value = int(self._settings.value("pdf_viewer/line_scale", 7, int) or 7)
        self._line_slider_value = max(1, min(10, self._line_slider_value))
        self._shape_value = str(self._settings.value("pdf_viewer/bubble_shape", "Circle") or "Circle")
        if self._shape_value not in ("Circle", "Rectangle"):
            self._shape_value = "Circle"

        # Bubble color (unselected)
        try:
            c = str(self._settings.value("pdf_viewer/bubble_color", "DC2828") or "DC2828").strip()
        except Exception:
            c = "DC2828"
        c = re.sub(r"[^0-9a-fA-F]", "", c)
        if len(c) != 6:
            c = "DC2828"
        self.bubble_color = QColor("#" + c)

        # Bubble backfill: when enabled, bubbles are filled solid color to mask drawing lines.
        try:
            self.bubble_backfill_white = bool(self._settings.value("pdf_viewer/bubble_backfill_white", True, bool))
        except Exception:
            self.bubble_backfill_white = True
        try:
            bf = str(self._settings.value("pdf_viewer/bubble_backfill_rgb", "FFFFFF") or "FFFFFF").strip()
        except Exception:
            bf = "FFFFFF"
        bf = re.sub(r"[^0-9a-fA-F]", "", bf)
        if len(bf) != 6:
            bf = "FFFFFF"
        self.bubble_backfill_rgb = bf.upper()

        # Whether to render PDF's own annotations (Kofax markup, etc.) in the background.
        try:
            self.show_pdf_annots = bool(self._settings.value("pdf_viewer/show_pdf_annots", True, bool))
        except Exception:
            self.show_pdf_annots = True
        
        # New granular visibility controls
        self.show_external_annots = True
        self.show_internal_annots = True
        
        # Whether to auto-import external annotations (overlay) if no internal ones exist.
        try:
            self.auto_import_annots = bool(self._settings.value("pdf_viewer/auto_import_annots", True, bool))
        except Exception:
            self.auto_import_annots = True

        self._pdf_annots_mutated: bool = False
        self._pdf_annots_deleted_all: bool = False
        self._pdf_annots_deleted_pages: set[int] = set()

        # Zoom limits (50% - 200%)
        self.MIN_ZOOM = 0.5
        self.MAX_ZOOM = 4.0

        # 100% baseline render scale used for good clarity
        # (we keep the previous tuned value so 100% looks the same as before)
        self.base_render_scale = 0.65
        self._zoom_factor = 1.0
        self._did_initial_fit = False

        # Bubble sizing
        # Keep 1-10 slider, but make the smallest usable size a bit smaller.
        self._bubble_radius_offset = 3  # was 5
        self._bubble_radius_step = 3

        # Image Enhancement
        self.enhance_mode = False
        self.brightness = 1.0
        self.contrast = 1.0
        self.sharpness = 1.0

        # Note-extraction regions (stored as normalized rects per page)
        # page_index -> list[(x0,y0,x1,y1)] with values in 0..1 relative to rendered page
        self.note_regions_by_page: dict[int, list[tuple[float, float, float, float]]] = {}
        self.note_region_items: list[QGraphicsRectItem] = []
        self.notes_dialog: NotesExtractDialog | None = None

        # Drawing grid overlay (not saved into PDF)
        self.grid_items: list[QGraphicsItem] = []
        self._grid_v_lines: list[QGraphicsLineItem] = []
        self._grid_h_lines: list[QGraphicsLineItem] = []
        self._grid_label_items: list[QGraphicsTextItem] = []
        self._grid_bounds_item: _GridBoundsItem | None = None
        try:
            self.grid_enabled = bool(self._settings.value("pdf_viewer/grid_enabled", False, bool))
        except Exception:
            self.grid_enabled = False
        try:
            self.grid_left_pct = float(self._settings.value("pdf_viewer/grid_left_pct", 0.0, float) or 0.0)
            self.grid_top_pct = float(self._settings.value("pdf_viewer/grid_top_pct", 0.0, float) or 0.0)
            self.grid_width_pct = float(self._settings.value("pdf_viewer/grid_width_pct", 100.0, float) or 100.0)
            self.grid_height_pct = float(self._settings.value("pdf_viewer/grid_height_pct", 100.0, float) or 100.0)
        except Exception:
            self.grid_left_pct = 0.0
            self.grid_top_pct = 0.0
            self.grid_width_pct = 100.0
            self.grid_height_pct = 100.0

        # Per-page rotation in degrees (0/90/180/270)
        self.page_rotation_by_page: dict[int, int] = {}
        self._last_render_rotation: int = 0
        
        # PDF state
        self.doc = None
        self.file_path = None
        self.current_page = 0
        self.total_pages = 0
        self._rendered_page_index: int | None = None
        # Use 72 DPI base (1:1 with PDF points) - let view scaling handle zoom
        # This prevents downscaling artifacts when zooming out
        self.base_dpi = 72
        self.current_render_scale = 1.0
        
        # Bubble state
        self.bubble_base_radius = self._radius_from_size_slider(self._size_slider_value)
        # Effective outline width (used for rendering/export). UI still shows 1-10.
        self.bubble_line_width = 3
        # Apply mapping to effective width.
        try:
            self.bubble_line_width = max(1, min(4, int(round(1.0 + ((self._line_slider_value - 1.0) * 3.0 / 9.0)))))
        except Exception:
            self.bubble_line_width = 3
        self.bubble_shape = self._shape_value
        # Scene items for the currently displayed page
        self.bubbles: list[BubbleItem] = []
        # page_index -> list[(start, end, x_norm, y_norm, base_radius)] stored in UNROTATED normalized coords
        self.bubble_specs_by_page: dict[int, list[tuple[int, int, float, float, int, str]]] = {}

        self.next_bubble_number = 1
        self.placing_mode = False
        self._range_end_number: int | None = None
        self.range_mode = False
        # Reference Location mode for Form 3: 'sheet_zone' | 'page_label' | 'none'
        self.reference_location_mode: str = "sheet_zone"
        
        self.page_items = []
        self.pixmap_item = None
        
        self._create_controls()
        self._setup_ui(embed_controls=self._embed_controls)
        self._setup_shortcuts()

        # Undo stack for bubble actions
        self._undo_stack: list[tuple[dict[int, list[tuple[int, int, float, float, int, str]]], int]] = []
        self._max_undo = 100

        # Track whether the user has modified bubbles since last save/load.
        self._dirty: bool = False
        self._last_saved_pdf_path: str | None = None
        self._opened_pdf_path: str | None = None
        # The PDF path whose directory hosts the sidecar JSON for editability.
        self._sidecar_context_pdf_path: str | None = None

    def _radius_from_size_slider(self, value: int) -> int:
        v = max(1, min(10, int(value)))
        # Map 1..10 -> old 2..10 (cap at 10 to keep the max unchanged).
        old_v = min(10, v + 1)
        return int(self._bubble_radius_offset + (old_v * self._bubble_radius_step))

    def _size_slider_from_radius(self, radius: int) -> int:
        try:
            r = int(radius)
        except Exception:
            return 5
        # Reverse mapping of _radius_from_size_slider.
        old_v = int(round((r - int(self._bubble_radius_offset)) / float(self._bubble_radius_step)))
        old_v = max(2, min(10, old_v))
        v = old_v - 1
        return max(1, min(10, int(v)))

    def _debug(self, msg: str) -> None:
        if not getattr(self, "_debug_enabled", False):
            return
        logger.debug("%s", msg)

    def is_dirty(self) -> bool:
        return bool(getattr(self, "_dirty", False))

    def _set_dirty(self, value: bool) -> None:
        self._dirty = bool(value)

    def _sidecar_path_for_pdf(self, pdf_path: str) -> str:
        return str(pdf_path) + ".as9102_bubbles.json"

    def _read_sidecar_json(self, pdf_path: str) -> dict | None:
        sidecar_path = self._sidecar_path_for_pdf(pdf_path)
        if not sidecar_path or not os.path.exists(sidecar_path):
            return None
        try:
            with open(sidecar_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                return data
        except Exception:
            return None
        return None

    def _load_edit_state_sidecar(self, pdf_path: str, data: dict | None = None) -> None:
        if data is None:
            data = self._read_sidecar_json(pdf_path) or {}
        if not isinstance(data, dict) or not data:
            return

        try:
            raw_specs = data.get("bubble_specs_by_page", {}) or {}
            specs_by_page: dict[int, list[tuple[int, int, float, float, int, str]]] = {}
            for k, v in raw_specs.items():
                page_index = int(k)
                out: list[tuple[int, int, float, float, int, str]] = []
                for item in (v or []):
                    try:
                        s, e, x, y, r = item[:5]
                        bf = item[5] if len(item) > 5 else ""
                    except Exception:
                        continue
                    out.append((int(s), int(e), float(x), float(y), int(r), str(bf or "")))
                specs_by_page[page_index] = out
            self.bubble_specs_by_page = specs_by_page

            raw_rot = data.get("page_rotation_by_page", {}) or {}
            rot_by_page: dict[int, int] = {}
            for k, v in raw_rot.items():
                try:
                    rot_by_page[int(k)] = int(v)
                except Exception:
                    pass
            if rot_by_page:
                self.page_rotation_by_page = rot_by_page

            if "next_bubble_number" in data:
                try:
                    self.next_bubble_number = int(data.get("next_bubble_number") or 1)
                except Exception:
                    self._recompute_next_bubble_number()
            else:
                self._recompute_next_bubble_number()
        except Exception:
            return

        self._set_dirty(False)
        self._last_saved_pdf_path = str(pdf_path)
        self._sidecar_context_pdf_path = str(pdf_path)

    def _save_edit_state_sidecar(self, pdf_path: str) -> None:
        if not pdf_path:
            return
        sidecar_path = self._sidecar_path_for_pdf(pdf_path)
        try:
            data = {
                "version": 1,
                # "pdf_path" is the PDF this sidecar lives next to (the file user opens).
                "pdf_path": str(pdf_path),
                # "source_pdf_path" is the clean/background PDF we rendered over when saving.
                "source_pdf_path": str(getattr(self, "file_path", "") or ""),
                "next_bubble_number": int(self.next_bubble_number),
                "bubble_specs_by_page": {
                    str(int(k)): [list(item) for item in (v or [])]
                    for k, v in (self.bubble_specs_by_page or {}).items()
                },
                "page_rotation_by_page": {
                    str(int(k)): int(v)
                    for k, v in (self.page_rotation_by_page or {}).items()
                },
            }
            with open(sidecar_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception:
            return

    def _extract_bubbles_from_doc_annotations(self, doc: "fitz.Document") -> dict[int, list[tuple[int, int, float, float, int, str]]]:
        """Best-effort import of bubble annotations from a PDF.

        The example "Formal Document" bubble set uses FreeText annotations.
        We treat FreeText contents like "12" or "3-5" as bubbles.
        """
        specs_by_page: dict[int, list[tuple[int, int, float, float, int, str]]] = {}
        if doc is None:
            return specs_by_page

        def _clean_content(s: str) -> str:
            s = str(s or "")
            # Normalize control characters like CR from some editors.
            s = re.sub(r"[\x00-\x1f\x7f]+", " ", s).strip()
            s = re.sub(r"\s+", " ", s).strip()
            return s

        def _parse_label_segments(s: str) -> list[tuple[int, int]]:
            """Parse an annotation label into one or more bubble segments.

            Supports:
            - "12" -> [(12,12)]
            - "3-5" -> [(3,5)]
            - "29-30-31" -> [(29,31)]
            - "1-2, 3, 5-7" -> [(1,2),(3,3),(5,7)]
            """
            s = _clean_content(s)
            if not s:
                return []

            # Reject obvious non-bubble text early (letters / decimals tend to be notes/dimensions).
            # Keep this conservative to avoid importing random numbers from notes.
            if re.search(r"[A-Za-z]", s):
                return []
            if "." in s:
                return []

            # Allow common wrappers without requiring the label to be ONLY digits/separators.
            # Example: "(12)", "12)", "#12", "12 - 14".
            s = re.sub(r"[\[\]\(\){}]", " ", s)
            s = re.sub(r"\s+", " ", s).strip()

            # Normalize various dash characters to '-'.
            s = (
                s.replace("\u2013", "-")
                .replace("\u2014", "-")
                .replace("\u2212", "-")
                .replace("\u2012", "-")
                .replace("\u2010", "-")
            )

            # Only treat annotations that look like bubble numbering.
            # (Prevents importing other annotation text that happens to include numbers.)
            if not re.match(r"^[0-9\s,;\-#]+$", s):
                return []

            parts = [p.strip() for p in re.split(r"[;,]", s) if p.strip()]
            if not parts:
                parts = [s]

            out: list[tuple[int, int]] = []

            for part in parts:
                nums = [int(m.group(0)) for m in re.finditer(r"\d{1,4}", part)]
                if not nums:
                    continue

                # If the part is a clean "a-b" range, treat it as a range.
                if (
                    len(nums) == 2
                    and re.match(r"^\s*\d{1,4}\s*-\s*\d{1,4}\s*$", part)
                ):
                    start, end = int(nums[0]), int(nums[1])
                    if end < start:
                        end = start
                    out.append((start, end))
                    continue

                # Otherwise treat it as a sequence of numbers and collapse consecutive runs.
                run_start = nums[0]
                prev = nums[0]
                for n in nums[1:]:
                    if n == prev + 1:
                        prev = n
                        continue
                    out.append((int(run_start), int(prev)))
                    run_start = n
                    prev = n
                out.append((int(run_start), int(prev)))

            # Special-case: some annotations use comma separation that effectively
            # continues numbering (e.g. "1-2,3" should become "1" and "2-3").
            # Rule: if a range (a-b) is immediately followed by a single (b+1),
            # shift the boundary number b into the right-hand segment.
            fixed: list[tuple[int, int]] = []
            i = 0
            while i < len(out):
                a, b = out[i]
                if (
                    i + 1 < len(out)
                    and b > a
                    and out[i + 1][0] == out[i + 1][1]
                    and out[i + 1][0] == b + 1
                ):
                    c = out[i + 1][0]
                    left_end = b - 1
                    if left_end >= a:
                        fixed.append((int(a), int(left_end)))
                    else:
                        fixed.append((int(a), int(a)))
                    fixed.append((int(b), int(c)))
                    i += 2
                    continue
                fixed.append((int(a), int(b)))
                i += 1

            return fixed

        def _annot_label(ann) -> str:
            info = getattr(ann, "info", {}) or {}
            # Prefer content; fall back to other common fields used by editors.
            for k in ("content", "contents", "Contents", "subject", "title", "name"):
                try:
                    v = info.get(k, "")
                except Exception:
                    v = ""
                if str(v or "").strip():
                    return str(v)
            return ""

        used_numbers: set[int] = set()

        # First pass: check for internal annotations (created by this app)
        has_internal_annots = False
        try:
            for page_index in range(int(getattr(doc, "page_count", 0) or 0)):
                try:
                    page = doc.load_page(page_index)
                    for ann in (page.annots() or []):
                        info = getattr(ann, "info", {}) or {}
                        if self._debug_enabled:
                            print(f"[AS9102_DEBUG_PDF] Page {page_index} Annot info: {info}", flush=True)

                        if info.get("title") == "AS9102_FAI_BUBBLE" or info.get("subject") == "AS9102_FAI_BUBBLE":
                            has_internal_annots = True
                            break
                except Exception:
                    pass
                if has_internal_annots:
                    break
        except Exception:
            pass

        if self._debug_enabled:
            print(f"[AS9102_DEBUG_PDF] _extract_bubbles: has_internal_annots={has_internal_annots} auto_import={self.auto_import_annots}")

        def _width_mult_for_label(label: str) -> float:
            n = len(str(label or ""))
            if n <= 2:
                return 1.0
            if n == 3:
                return 1.20
            if n == 4:
                return 1.35
            if n == 5:
                return 1.55
            if n == 6:
                return 1.75
            if n == 7:
                return 1.95
            if n <= 9:
                return 2.20
            return 2.50

        for page_index in range(int(getattr(doc, "page_count", 0) or 0)):
            try:
                page = doc.load_page(page_index)
            except Exception:
                continue

            page_rect = getattr(page, "rect", None)
            if page_rect is None or page_rect.width <= 0 or page_rect.height <= 0:
                continue

            try:
                annots = list(page.annots() or [])
            except Exception:
                annots = []

            # Pre-scan for internal bubbles on this page to avoid collisions
            internal_centers: list[tuple[float, float]] = []
            if has_internal_annots:
                for ann in annots:
                    try:
                        info = getattr(ann, "info", {}) or {}
                        if info.get("title") == "AS9102_FAI_BUBBLE" or info.get("subject") == "AS9102_FAI_BUBBLE":
                            r = ann.rect
                            cx = (r.x0 + r.x1) / 2.0
                            cy = (r.y0 + r.y1) / 2.0
                            internal_centers.append((float(cx), float(cy)))
                    except Exception:
                        pass

            items: list[tuple[int, int, float, float, int, str]] = []

            for ann in annots:
                try:
                    info = getattr(ann, "info", {}) or {}
                    is_internal = (info.get("title") == "AS9102_FAI_BUBBLE" or info.get("subject") == "AS9102_FAI_BUBBLE")

                    # Logic:
                    # 1. Always import internal bubbles.
                    # 2. If external:
                    #    - Check auto_import_annots.
                    #    - Check for collision with ANY internal bubble on this page.
                    #      If it collides, we assume it's the "underlying" external bubble for an existing internal one, so skip it.
                    
                    if is_internal:
                        pass
                    else:
                        if not self.auto_import_annots:
                            continue
                        
                        # Check collision with internal bubbles
                        is_colliding = False
                        try:
                            r = ann.rect
                            acx = (r.x0 + r.x1) / 2.0
                            acy = (r.y0 + r.y1) / 2.0
                            for (icx, icy) in internal_centers:
                                if abs(acx - icx) < 15.0 and abs(acy - icy) < 15.0:
                                    is_colliding = True
                                    break
                        except Exception:
                            pass
                        
                        if is_colliding:
                            continue

                    # Kofax and other tools may store these as various annot types.
                    label = _annot_label(ann)
                    segments = _parse_label_segments(label)
                    if not segments:
                        continue
                    rect = ann.rect
                    cx = (float(rect.x0) + float(rect.x1)) / 2.0
                    cy = (float(rect.y0) + float(rect.y1)) / 2.0
                    rx = cx / float(page_rect.width)
                    ry = cy / float(page_rect.height)
                    rx = max(0.0, min(1.0, float(rx)))
                    ry = max(0.0, min(1.0, float(ry)))

                    # Use current bubble size setting so imported bubbles match the UI.
                    try:
                        br = int(getattr(self, "bubble_base_radius", 15) or 15)
                    except Exception:
                        br = 15
                    br = max(5, min(80, int(br)))

                    # Place multiple derived bubbles side-by-side without overlapping.
                    seg_count = len(segments)
                    if seg_count <= 0:
                        continue

                    # Compute spacing from bubble width (scene units -> PDF points).
                    pad_scene = max(6.0, float(br) * 0.6)
                    # Default spacing (will be refined per-label below); keep stable center.
                    spacing_pts_default = (2.0 * float(br) + pad_scene) / float(self.base_render_scale)

                    # Precompute widths per segment label.
                    labels: list[str] = []
                    for (s0, e0) in segments:
                        try:
                            s0i = int(s0)
                            e0i = int(e0)
                        except Exception:
                            labels.append("")
                            continue
                        labels.append(f"{s0i}-{e0i}" if e0i > s0i else str(s0i))

                    half_ws_pts: list[float] = []
                    for lab in labels:
                        half_w_scene = float(br) * float(_width_mult_for_label(lab))
                        half_ws_pts.append(half_w_scene / float(self.base_render_scale))

                    # Compute centers so items don't overlap.
                    centers_pts: list[float] = []
                    cur_x = float(cx)
                    # Build positions relative to cx; first compute total width.
                    total_w = 0.0
                    for i in range(seg_count):
                        hw = half_ws_pts[i] if i < len(half_ws_pts) else (float(br) / float(self.base_render_scale))
                        total_w += 2.0 * float(hw)
                        if i != seg_count - 1:
                            total_w += (pad_scene / float(self.base_render_scale))
                    left = float(cx) - total_w / 2.0
                    x_cursor = left
                    for i in range(seg_count):
                        hw = half_ws_pts[i] if i < len(half_ws_pts) else (float(br) / float(self.base_render_scale))
                        centers_pts.append(x_cursor + float(hw))
                        x_cursor += 2.0 * float(hw)
                        if i != seg_count - 1:
                            x_cursor += (pad_scene / float(self.base_render_scale))

                    for i, (start, end) in enumerate(segments):
                        start = int(start)
                        end = int(end)
                        if end < start:
                            end = start
                        if start <= 0:
                            continue
                        if end - start > 9999:
                            end = start

                        # Skip duplicates (any overlap) to match current behavior.
                        overlap = any((n in used_numbers) for n in range(start, end + 1))
                        if overlap:
                            continue

                        try:
                            cx_i = float(centers_pts[i])
                        except Exception:
                            cx_i = float(cx) + float(i) * float(spacing_pts_default)
                        rx_i = max(0.0, min(1.0, float(cx_i) / float(page_rect.width)))
                        ry_i = ry

                        items.append((int(start), int(end), float(rx_i), float(ry_i), int(br), ""))
                        for n in range(start, end + 1):
                            used_numbers.add(int(n))
                except Exception:
                    continue

            if getattr(self, "_debug_enabled", False):
                try:
                    print(
                        f"[AS9102_DEBUG_PDF]   page={page_index+1} annots={len(annots)} imported_specs={len(items)}",
                        flush=True,
                    )
                except Exception:
                    pass

            if items:
                items.sort(key=lambda t: (t[0], t[1], t[2], t[3]))
                specs_by_page[int(page_index)] = items

        return specs_by_page

    def can_close(self) -> bool:
        if not self.is_dirty():
            return True

        mb = QMessageBox(self)
        mb.setIcon(QMessageBox.Question)
        mb.setWindowTitle("Unsaved Changes")
        mb.setText("Save changes to the drawing before closing?")
        mb.setStandardButtons(QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel)
        mb.setDefaultButton(QMessageBox.Save)
        choice = mb.exec()

        if choice == QMessageBox.Save:
            return bool(self.save_drawing())
        if choice == QMessageBox.Discard:
            return True
        return False

    def _create_controls(self) -> None:
        # Bubble controls
        self.add_bubble_btn = QPushButton("Add Bubble")
        self.add_bubble_btn.setCheckable(True)
        self.add_bubble_btn.clicked.connect(self.toggle_placing_mode)

        self.bubble_number_spin = QSpinBox()
        self.bubble_number_spin.setRange(1, 9999)
        self.bubble_number_spin.setValue(int(getattr(self, "next_bubble_number", 1) or 1))
        self.bubble_number_spin.setEnabled(True)
        try:
            self.bubble_number_spin.setMaximumWidth(90)
        except Exception:
            pass
        self.bubble_number_spin.valueChanged.connect(self._on_pending_bubble_number_changed)

        self.add_range_btn = QPushButton("Add Range")
        self.add_range_btn.setCheckable(True)
        self.add_range_btn.toggled.connect(self.toggle_range_mode)

        self.clear_btn = QPushButton("Clear All")
        self.clear_btn.clicked.connect(self.clear_bubbles)

        # Bubble size (1-10)
        self.size_slider = QSlider(Qt.Horizontal)
        self.size_slider.setMinimum(1)
        self.size_slider.setMaximum(10)
        self.size_slider.setValue(int(getattr(self, "_size_slider_value", 5) or 5))
        self.size_slider.setMaximumWidth(160)
        self.size_slider.valueChanged.connect(self.on_size_changed)

        self.size_label = QLabel(str(int(getattr(self, "_size_slider_value", 5) or 5)))
        self.size_label.setMinimumWidth(20)

        self.shape_combo = QComboBox()
        self.shape_combo.addItems(["Circle", "Rectangle"])
        self.shape_combo.setCurrentText(self.bubble_shape)
        self.shape_combo.currentIndexChanged.connect(self.on_shape_changed)

        self.line_slider = QSlider(Qt.Horizontal)
        self.line_slider.setMinimum(1)
        self.line_slider.setMaximum(10)
        self.line_slider.setValue(int(getattr(self, "_line_slider_value", 7) or 7))
        self.line_slider.setMaximumWidth(160)
        self.line_slider.valueChanged.connect(self.on_line_width_changed)

        # Enhancement Controls
        self.enhance_btn = QPushButton("Enhance Image")
        self.enhance_btn.setCheckable(True)
        self.enhance_btn.clicked.connect(self.toggle_enhance_mode)

        self.enhance_panel = QWidget()
        self.enhance_panel.setVisible(False)
        enhance_layout = QFormLayout(self.enhance_panel)
        enhance_layout.setContentsMargins(5, 5, 5, 5)
        try:
            enhance_layout.setLabelAlignment(Qt.AlignLeft)
        except Exception:
            pass

        # Brightness
        self.brightness_slider = QSlider(Qt.Horizontal)
        self.brightness_slider.setRange(1, 30)  # 0.1 to 3.0
        self.brightness_slider.setValue(10)
        try:
            self.brightness_slider.setMaximumWidth(220)
        except Exception:
            pass
        self.brightness_slider.valueChanged.connect(self.on_enhancement_changed)
        enhance_layout.addRow("Brightness", self.brightness_slider)

        # Contrast
        self.contrast_slider = QSlider(Qt.Horizontal)
        self.contrast_slider.setRange(1, 30)
        self.contrast_slider.setValue(10)
        try:
            self.contrast_slider.setMaximumWidth(220)
        except Exception:
            pass
        self.contrast_slider.valueChanged.connect(self.on_enhancement_changed)
        enhance_layout.addRow("Contrast", self.contrast_slider)

        # Sharpness
        self.sharpness_slider = QSlider(Qt.Horizontal)
        self.sharpness_slider.setRange(1, 30)
        self.sharpness_slider.setValue(10)
        try:
            self.sharpness_slider.setMaximumWidth(220)
        except Exception:
            pass
        self.sharpness_slider.valueChanged.connect(self.on_enhancement_changed)
        enhance_layout.addRow("Sharpness", self.sharpness_slider)

        # Reset
        self.reset_enhance_btn = QPushButton("Reset")
        self.reset_enhance_btn.clicked.connect(self.reset_enhancements)
        enhance_layout.addRow("", self.reset_enhance_btn)

        # Show the UI scale (1-10), not the internal mapped thickness.
        self.line_label = QLabel(str(int(getattr(self, "_line_slider_value", 7) or 7)))
        self.line_label.setMinimumWidth(20)

        # Notes
        self.add_note_region_btn = QPushButton("Notes Window")
        self.add_note_region_btn.setCheckable(True)
        self.add_note_region_btn.clicked.connect(self.toggle_note_region_mode)

        self.clear_note_regions_btn = QPushButton("Clear Notes Window")
        self.clear_note_regions_btn.clicked.connect(self.clear_note_regions)

        self.notes_mode_combo = QComboBox()
        self.notes_mode_combo.addItems(["Auto", "PDF Text", "OCR"])
        self.notes_mode_combo.setMinimumWidth(90)

        self.extract_notes_btn = QPushButton("Extract Notes")
        self.extract_notes_btn.clicked.connect(self.extract_notes)

        # Traditional Save/Save As behavior.
        self.save_drawing_btn = QPushButton("Save")
        self.save_drawing_btn.clicked.connect(self.save_drawing)

        self.save_drawing_as_btn = QPushButton("Save As")
        self.save_drawing_as_btn.clicked.connect(self.save_drawing_as)

        # Selection
        self.select_page_btn = QPushButton("Select Page")
        self.select_page_btn.clicked.connect(self.select_page_bubbles)

        self.select_all_btn = QPushButton("Select All")
        self.select_all_btn.clicked.connect(self.select_all_bubbles)

        # Clipboard
        self.copy_btn = QPushButton("Copy")
        self.copy_btn.clicked.connect(self.copy_bubbles)

        self.paste_btn = QPushButton("Paste")
        self.paste_btn.clicked.connect(self.paste_bubbles)

        # Navigation
        self.page_combo = QComboBox()
        self.page_combo.setMinimumWidth(90)
        self.page_combo.currentIndexChanged.connect(self.go_to_page)

        self.prev_btn = QPushButton("◀ Prev")
        self.prev_btn.clicked.connect(self.prev_page)

        self.next_btn = QPushButton("Next ▶")
        self.next_btn.clicked.connect(self.next_page)

        # Zoom
        self.zoom_out_btn = QPushButton("−")
        self.zoom_out_btn.setMaximumWidth(30)
        self.zoom_out_btn.clicked.connect(self.zoom_out)

        self.zoom_fit_btn = QPushButton("Fit")
        self.zoom_fit_btn.clicked.connect(self.fit_to_view)

        self.zoom_100_btn = QPushButton("100%")
        self.zoom_100_btn.clicked.connect(self.zoom_100)

        self.zoom_in_btn = QPushButton("+")
        self.zoom_in_btn.setMaximumWidth(30)
        self.zoom_in_btn.clicked.connect(self.zoom_in)

        self.zoom_label = QLabel("100%")
        self.zoom_label.setMinimumWidth(55)

        # Rotate
        self.rotate_left_btn = QPushButton("⟲")
        self.rotate_left_btn.setToolTip("Rotate Left")
        self.rotate_left_btn.clicked.connect(self.rotate_left)

        self.rotate_right_btn = QPushButton("⟳")
        self.rotate_right_btn.setToolTip("Rotate Right")
        self.rotate_right_btn.clicked.connect(self.rotate_right)

    def _setup_ui(self, *, embed_controls: bool) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)

        if embed_controls:
            # Top toolbar (bubble + notes)
            toolbar = QHBoxLayout()
            toolbar.setSpacing(5)

            toolbar.addWidget(self.add_bubble_btn)
            toolbar.addWidget(self.bubble_number_spin)
            toolbar.addWidget(self.add_range_btn)
            toolbar.addWidget(self.clear_btn)
            toolbar.addWidget(QLabel(" | Size:"))
            toolbar.addWidget(self.size_slider)
            toolbar.addWidget(self.size_label)
            toolbar.addWidget(QLabel(" | Shape:"))
            toolbar.addWidget(self.shape_combo)
            toolbar.addWidget(QLabel(" | Line:"))
            toolbar.addWidget(self.line_slider)
            toolbar.addWidget(self.line_label)
            toolbar.addWidget(QLabel(" | Notes:"))
            toolbar.addWidget(self.notes_mode_combo)
            toolbar.addWidget(self.add_note_region_btn)
            toolbar.addWidget(self.clear_note_regions_btn)
            toolbar.addWidget(self.extract_notes_btn)
            toolbar.addWidget(self.save_drawing_btn)
            toolbar.addWidget(self.save_drawing_as_btn)
            toolbar.addWidget(QLabel(" | Select:"))
            toolbar.addWidget(self.select_page_btn)
            toolbar.addWidget(self.select_all_btn)
            toolbar.addWidget(QLabel(" | Clipboard:"))
            toolbar.addWidget(self.copy_btn)
            toolbar.addWidget(self.paste_btn)
            toolbar.addWidget(QLabel(" | "))
            toolbar.addWidget(self.enhance_btn)
            toolbar.addStretch()

            toolbar_widget = QWidget()
            toolbar_widget.setLayout(toolbar)
            layout.addWidget(toolbar_widget)
            
            layout.addWidget(self.enhance_panel)
        
        # Graphics view
        self.scene = QGraphicsScene()
        self.view = InteractiveGraphicsView(self.scene, self)
        self.view.bubble_click.connect(self.on_bubble_click)
        self.view.note_region_created.connect(self.on_note_region_created)
        layout.addWidget(self.view)

        if embed_controls:
            # Bottom bar (page + zoom + rotate), centered
            bottom_bar = QHBoxLayout()
            bottom_bar.setSpacing(6)
            bottom_bar.addStretch()

            bottom_bar.addWidget(QLabel("Page:"))
            bottom_bar.addWidget(self.page_combo)
            bottom_bar.addWidget(self.prev_btn)
            bottom_bar.addWidget(self.next_btn)
            bottom_bar.addWidget(QLabel(" | Zoom:"))
            bottom_bar.addWidget(self.zoom_out_btn)
            bottom_bar.addWidget(self.zoom_fit_btn)
            bottom_bar.addWidget(self.zoom_100_btn)
            bottom_bar.addWidget(self.zoom_in_btn)
            bottom_bar.addWidget(self.zoom_label)
            bottom_bar.addWidget(QLabel(" | Rotate:"))
            bottom_bar.addWidget(self.rotate_left_btn)
            bottom_bar.addWidget(self.rotate_right_btn)
            bottom_bar.addStretch()

            bottom_widget = QWidget()
            bottom_widget.setLayout(bottom_bar)
            layout.addWidget(bottom_widget)

    def _setup_shortcuts(self):
        """Setup keyboard shortcuts."""
        self._undo_shortcut = QShortcut(QKeySequence("Ctrl+Z"), self)
        self._undo_shortcut.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
        self._undo_shortcut.activated.connect(lambda: (self._debug("Ctrl+Z activated"), self.undo_last_action()))

        self._delete_shortcut = QShortcut(QKeySequence(Qt.Key_Delete), self)
        self._delete_shortcut.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
        self._delete_shortcut.activated.connect(self.delete_selected_bubbles)

        self._backspace_shortcut = QShortcut(QKeySequence(Qt.Key_Backspace), self)
        self._backspace_shortcut.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
        self._backspace_shortcut.activated.connect(self.delete_selected_bubbles)

        self._esc_shortcut = QShortcut(QKeySequence(Qt.Key_Escape), self)
        self._esc_shortcut.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
        self._esc_shortcut.activated.connect(self.on_escape)

        self._select_page_shortcut = QShortcut(QKeySequence("Ctrl+A"), self)
        self._select_page_shortcut.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
        self._select_page_shortcut.activated.connect(self.select_page_bubbles)

        self._select_all_shortcut = QShortcut(QKeySequence("Ctrl+Shift+A"), self)
        self._select_all_shortcut.setContext(Qt.ApplicationShortcut)
        self._select_all_shortcut.activated.connect(self.select_all_bubbles)

        self._copy_shortcut = QShortcut(QKeySequence("Ctrl+C"), self)
        self._copy_shortcut.setContext(Qt.ApplicationShortcut)
        self._copy_shortcut.activated.connect(self.copy_bubbles)

        self._copy_all_shortcut = QShortcut(QKeySequence("Ctrl+Shift+C"), self)
        self._copy_all_shortcut.setContext(Qt.ApplicationShortcut)
        self._copy_all_shortcut.activated.connect(lambda: self.copy_bubbles(copy_all_pages=True))

        self._paste_shortcut = QShortcut(QKeySequence("Ctrl+V"), self)
        self._paste_shortcut.setContext(Qt.ApplicationShortcut)
        self._paste_shortcut.activated.connect(self.paste_bubbles)

    def delete_selected_bubbles(self) -> None:
        """Delete any selected BubbleItem(s) on the current page."""
        if getattr(self, "scene", None) is None:
            return
        try:
            selected = list(self.scene.selectedItems() or [])
        except Exception:
            selected = []
        bubbles_to_delete: list[BubbleItem] = []
        for it in selected:
            if isinstance(it, BubbleItem):
                bubbles_to_delete.append(it)
        if not bubbles_to_delete:
            return

        self._push_undo_state()
        changed = False
        for bubble in bubbles_to_delete:
            if bubble in self.bubbles:
                try:
                    self.bubbles.remove(bubble)
                except Exception:
                    pass
                try:
                    self.scene.removeItem(bubble)
                except Exception:
                    pass
                try:
                    self.bubble_removed.emit(int(getattr(bubble, "number", 0) or 0))
                except Exception:
                    pass
                changed = True

        if changed:
            try:
                self._persist_current_page_bubbles(rotation=self._last_render_rotation)
            except Exception:
                pass
            self._set_dirty(True)

    def _bubble_to_unrotated_spec(self, bubble: BubbleItem) -> tuple[int, int, float, float, int, str] | None:
        if self.pixmap_item is None:
            return None
        page_rect = self.pixmap_item.boundingRect()
        if page_rect.width() <= 0 or page_rect.height() <= 0:
            return None
        try:
            rx = float(bubble.pos().x() / page_rect.width())
            ry = float(bubble.pos().y() / page_rect.height())
        except Exception:
            return None
        rx = max(0.0, min(1.0, rx))
        ry = max(0.0, min(1.0, ry))

        rot = int(self._last_render_rotation or 0) % 360
        if rot:
            inv = (-rot) % 360
            rx, ry = self._rotate_norm_point(rx, ry, inv)
        try:
            start = int(getattr(bubble, "number", 0) or 0)
            end = int(getattr(bubble, "range_end", start) or start)
            br = int(getattr(bubble, "base_radius", self.bubble_base_radius) or self.bubble_base_radius)
        except Exception:
            return None
        try:
            bf = getattr(bubble, "backfill_rgb", "")
        except Exception:
            bf = ""
        return (start, end, float(rx), float(ry), int(br), str(bf or ""))

    def select_page_bubbles(self) -> None:
        # Select all bubbles on the current page (scene).
        try:
            for it in self.scene.selectedItems():
                try:
                    it.setSelected(False)
                except Exception:
                    pass
        except Exception:
            pass
        for b in list(self.bubbles or []):
            try:
                b.setSelected(True)
            except Exception:
                pass

    def select_all_bubbles(self) -> None:
        # There is only one page's bubbles in the scene at a time; treat "Select All" as select page.
        self.select_page_bubbles()

    def _clipboard_set(self, payload: dict) -> None:
        try:
            md = QMimeData()
            md.setData("application/x-as9102-bubbles+json", json.dumps(payload).encode("utf-8"))
            # Also put a text form for easy debugging.
            md.setText(json.dumps(payload))
            QApplication.clipboard().setMimeData(md)
        except Exception:
            try:
                QApplication.clipboard().setText(json.dumps(payload))
            except Exception:
                pass

    def _clipboard_get(self) -> dict | None:
        cb = QApplication.clipboard()
        md = cb.mimeData() if cb is not None else None
        if md is None:
            return None
        raw = None
        try:
            if md.hasFormat("application/x-as9102-bubbles+json"):
                raw = bytes(md.data("application/x-as9102-bubbles+json")).decode("utf-8", errors="ignore")
        except Exception:
            raw = None
        if not raw:
            try:
                raw = md.text()
            except Exception:
                raw = None
        if not raw:
            return None
        try:
            data = json.loads(raw)
            if isinstance(data, dict) and data.get("kind") == "as9102_bubbles":
                return data
        except Exception:
            return None
        return None

    def copy_bubbles(self, copy_all_pages: bool = False) -> None:
        if not self.doc:
            return
        try:
            self._persist_current_page_bubbles(rotation=self._last_render_rotation)
        except Exception:
            pass

        # If any bubbles are selected, copy selection by default.
        selected_bubbles: list[BubbleItem] = []
        try:
            for it in self.scene.selectedItems():
                if isinstance(it, BubbleItem):
                    selected_bubbles.append(it)
        except Exception:
            selected_bubbles = []

        if copy_all_pages:
            payload = {
                "kind": "as9102_bubbles",
                "mode": "all_pages",
                "bubble_specs_by_page": {
                    str(int(k)): [list(item) for item in (v or [])]
                    for k, v in (self.bubble_specs_by_page or {}).items()
                },
            }
            self._clipboard_set(payload)
            return

        if selected_bubbles:
            specs: list[list] = []
            for b in selected_bubbles:
                spec = self._bubble_to_unrotated_spec(b)
                if spec is not None:
                    specs.append(list(spec))
            payload = {
                "kind": "as9102_bubbles",
                "mode": "selection",
                "source_page": int(self.current_page),
                "bubbles": specs,
            }
            self._clipboard_set(payload)
            return

        # Otherwise copy current page.
        payload = {
            "kind": "as9102_bubbles",
            "mode": "page",
            "source_page": int(self.current_page),
            "bubbles": [list(item) for item in (self.bubble_specs_by_page.get(self.current_page, []) or [])],
        }
        self._clipboard_set(payload)

    def paste_bubbles(self) -> None:
        if not self.doc:
            return
        data = self._clipboard_get()
        if not data:
            return

        mode = str(data.get("mode") or "")

        # Pasting changes state
        self._push_undo_state()

        if mode == "all_pages":
            raw = data.get("bubble_specs_by_page", {}) or {}
            for k, v in raw.items():
                try:
                    page_index = int(k)
                except Exception:
                    continue
                if page_index < 0 or page_index >= int(self.total_pages or 0):
                    continue
                incoming: list[tuple[int, int, float, float, int, str]] = []
                try:
                    for item in (v or []):
                        try:
                            s, e, x, y, r = item[:5]
                            bf = item[5] if len(item) > 5 else ""
                        except Exception:
                            continue
                        incoming.append((int(s), int(e), float(x), float(y), int(r), str(bf or "")))
                except Exception:
                    continue
                existing = list(self.bubble_specs_by_page.get(page_index, []) or [])
                existing.extend(incoming)
                existing.sort(key=lambda t: (t[0], t[1], t[2], t[3]))
                self.bubble_specs_by_page[page_index] = existing

            self._recompute_next_bubble_number()
            self._rendered_page_index = None
            self.render_current_page(target_scale=self.current_render_scale, source_rotation=self._last_render_rotation)
            self._set_dirty(True)
            return

        if mode in ("selection", "page"):
            raw_bubbles = data.get("bubbles", []) or []
            incoming: list[tuple[int, int, float, float, int, str]] = []
            for item in raw_bubbles:
                try:
                    s, e, x, y, r = item[:5]
                    bf = item[5] if len(item) > 5 else ""
                    incoming.append((int(s), int(e), float(x), float(y), int(r), str(bf or "")))
                except Exception:
                    continue

            if not incoming:
                return

            existing = list(self.bubble_specs_by_page.get(self.current_page, []) or [])
            existing.extend(incoming)
            existing.sort(key=lambda t: (t[0], t[1], t[2], t[3]))
            self.bubble_specs_by_page[self.current_page] = existing
            self._recompute_next_bubble_number()

            self._rendered_page_index = None
            self.render_current_page(target_scale=self.current_render_scale, source_rotation=self._last_render_rotation)
            self._set_dirty(True)
            return

    def _push_undo_state(self) -> None:
        """Snapshot bubble state for Ctrl+Z."""
        self._debug("_push_undo_state")
        try:
            # Ensure we capture latest positions first.
            self._persist_current_page_bubbles(rotation=self._last_render_rotation)
        except Exception:
            pass

        snapshot: dict[int, list[tuple[int, int, float, float, int, str]]] = {}
        for page_index, specs in (self.bubble_specs_by_page or {}).items():
            out: list[tuple[int, int, float, float, int, str]] = []
            for spec in list(specs):
                try:
                    s, e, x, y, r = spec[:5]
                    bf = spec[5] if len(spec) > 5 else ""
                except Exception:
                    continue
                out.append((int(s), int(e), float(x), float(y), int(r), str(bf or "")))
            snapshot[int(page_index)] = out
        self._undo_stack.append((snapshot, int(self.next_bubble_number)))
        if len(self._undo_stack) > int(self._max_undo):
            self._undo_stack = self._undo_stack[-int(self._max_undo):]

    def undo_last_action(self) -> None:
        self._debug(f"undo_last_action: stack_size={len(self._undo_stack)}")
        if not self._undo_stack:
            return
        snapshot, next_num = self._undo_stack.pop()
        self.bubble_specs_by_page = {int(k): list(v) for k, v in snapshot.items()}
        self.next_bubble_number = int(next_num)

        # Prevent render_current_page() from persisting the *current* scene state
        # back into bubble_specs_by_page (which would overwrite the snapshot we
        # just restored).
        self._rendered_page_index = None

        # Exit placement modes to avoid confusion
        self.placing_mode = False
        self.range_mode = False
        self._range_end_number = None
        try:
            self.add_bubble_btn.setChecked(False)
            self.add_bubble_btn.setText("Add Bubble")
            self.add_range_btn.setChecked(False)
        except Exception:
            pass
        try:
            self.view.set_placing_mode(False)
        except Exception:
            pass

        # Re-render current page with current zoom
        self.render_current_page(target_scale=self.current_render_scale, source_rotation=self._last_render_rotation)

        self._set_dirty(True)
        
    def on_escape(self):
        if self.placing_mode:
            self.toggle_placing_mode()

    def _current_page_rotation(self) -> int:
        rot = int(self.page_rotation_by_page.get(self.current_page, 0) or 0)
        rot = rot % 360
        if rot not in (0, 90, 180, 270):
            # Snap to nearest right angle
            rot = int(round(rot / 90.0) * 90) % 360
        return rot

    def _rotate_norm_point(self, x: float, y: float, delta_degrees: int) -> tuple[float, float]:
        d = int(delta_degrees) % 360
        if d == 90:
            return (1.0 - y, x)
        if d == 180:
            return (1.0 - x, 1.0 - y)
        if d == 270:
            return (y, 1.0 - x)
        return (x, y)

    def _rotate_norm_rect(self, x0: float, y0: float, x1: float, y1: float, delta_degrees: int) -> tuple[float, float, float, float]:
        pts = [
            self._rotate_norm_point(x0, y0, delta_degrees),
            self._rotate_norm_point(x1, y0, delta_degrees),
            self._rotate_norm_point(x0, y1, delta_degrees),
            self._rotate_norm_point(x1, y1, delta_degrees),
        ]
        xs = [p[0] for p in pts]
        ys = [p[1] for p in pts]
        nx0 = max(0.0, min(1.0, min(xs)))
        nx1 = max(0.0, min(1.0, max(xs)))
        ny0 = max(0.0, min(1.0, min(ys)))
        ny1 = max(0.0, min(1.0, max(ys)))
        return (nx0, ny0, nx1, ny1)

    def rotate_left(self) -> None:
        if not self.doc:
            return
        old_rot = self._last_render_rotation
        self._persist_current_page_bubbles(rotation=old_rot)
        new_rot = (self._current_page_rotation() - 90) % 360
        self.page_rotation_by_page[self.current_page] = new_rot

        self._set_dirty(True)

        self.render_current_page(target_scale=self.current_render_scale, preserve_center=True, source_rotation=old_rot)

    def rotate_right(self) -> None:
        if not self.doc:
            return
        old_rot = self._last_render_rotation
        self._persist_current_page_bubbles(rotation=old_rot)
        new_rot = (self._current_page_rotation() + 90) % 360
        self.page_rotation_by_page[self.current_page] = new_rot

        self._set_dirty(True)

        self.render_current_page(target_scale=self.current_render_scale, preserve_center=True, source_rotation=old_rot)
        
    def on_size_changed(self, value):
        """Handle bubble size change.

        Slider range is 1-10, but value=1 maps to the previous "size 2" so bubbles
        never get too small to read.
        """
        try:
            self._size_slider_value = max(1, min(10, int(value)))
            self._settings.setValue("pdf_viewer/bubble_size", int(self._size_slider_value))
        except Exception:
            pass

        v = max(1, min(10, int(value)))
        self.bubble_base_radius = self._radius_from_size_slider(v)
        self.size_label.setText(str(v))
        
        for bubble in self.bubbles:
            bubble.set_base_radius(self.bubble_base_radius)

        try:
            self._persist_current_page_bubbles(rotation=self._last_render_rotation)
        except Exception:
            pass
        self._set_dirty(True)

    def on_shape_changed(self, _index: int) -> None:
        self.bubble_shape = self.shape_combo.currentText() or "Circle"
        try:
            self._settings.setValue("pdf_viewer/bubble_shape", str(self.bubble_shape))
        except Exception:
            pass
        for bubble in self.bubbles:
            bubble.update()

    def on_line_width_changed(self, value: int) -> None:
        v = max(1, min(10, int(value)))
        self._line_slider_value = v
        try:
            self._settings.setValue("pdf_viewer/line_scale", int(v))
        except Exception:
            pass
        # Map 1..10 UI scale to effective 1..4 thickness.
        # Requirement: UI=10 should match the old thickness 4.
        eff = int(round(1.0 + ((v - 1.0) * 3.0 / 9.0)))
        self.bubble_line_width = max(1, min(4, eff))
        self.line_label.setText(str(v))
        # BubbleItem.boundingRect() depends on the current line width.
        # When it changes, Qt requires prepareGeometryChange() or repaints/clipping can glitch
        # (notably: text may appear to disappear at max thickness).
        for bubble in self.bubbles:
            try:
                bubble.prepareGeometryChange()
            except Exception:
                pass
            bubble.update()

        try:
            if self.scene is not None:
                self.scene.update()
        except Exception:
            pass

    def toggle_placing_mode(self):
        """Toggle bubble placement mode."""
        if getattr(self, "add_range_btn", None) is not None and self.add_range_btn.isChecked():
            self.add_range_btn.setChecked(False)
        self.placing_mode = not self.placing_mode
        if not self.placing_mode:
            self._range_end_number = None
        if self.placing_mode:
            self.set_note_region_mode(False)
        self.add_bubble_btn.setChecked(self.placing_mode)
        self.view.set_placing_mode(self.placing_mode)

        if self.placing_mode:
            try:
                self._set_pending_bubble_number(self._lowest_available_number())
            except Exception:
                self._set_pending_bubble_number(int(self.next_bubble_number))
            self.add_bubble_btn.setText("Click to place")
        else:
            self.add_bubble_btn.setText("Add Bubble")

    def _existing_bubbled_numbers(self) -> set[int]:
        try:
            # Ensure current page state is persisted before checking.
            self._persist_current_page_bubbles(rotation=self._last_render_rotation, recompute_next=False)
        except Exception:
            pass
        try:
            return set(self.get_bubbled_numbers() or set())
        except Exception:
            return set()

    def _range_overlap(self, start: int, end: int) -> list[int]:
        start = int(start)
        end = int(end)
        if end < start:
            end = start
        existing = self._existing_bubbled_numbers()
        overlap = [n for n in range(start, end + 1) if n in existing]
        return overlap

    def _pending_bubble_number(self) -> int:
        try:
            return int(self.bubble_number_spin.value())
        except Exception:
            return int(getattr(self, "next_bubble_number", 1) or 1)

    def _set_pending_bubble_number(self, n: int) -> None:
        try:
            v = max(1, min(9999, int(n)))
        except Exception:
            v = 1
        try:
            self.bubble_number_spin.blockSignals(True)
            self.bubble_number_spin.setValue(int(v))
        except Exception:
            pass
        finally:
            try:
                self.bubble_number_spin.blockSignals(False)
            except Exception:
                pass

    def _next_available_number(self, start: int) -> int:
        # Backward compat: existing callers expect "a free number".
        # New behavior requested: always use the lowest available bubble number.
        return int(self._lowest_available_number())

    def _lowest_available_number(self) -> int:
        existing = self._existing_bubbled_numbers()
        n = 1
        while n in existing and n < 9999:
            n += 1
        return int(n)

    def _on_pending_bubble_number_changed(self, value: int) -> None:
        # If the user selects a bubble number that already exists, show an error
        # and move them to the next available number.
        if bool(getattr(self, "_pending_number_adjusting", False)):
            return

        try:
            n = int(value)
        except Exception:
            return

        overlap = self._range_overlap(n, n)
        if not overlap:
            return

        fixed = self._lowest_available_number()
        try:
            self._pending_number_adjusting = True
            QMessageBox.warning(
                self,
                "Duplicate Bubble Number",
                f"Bubble number {n} already exists. Next available is {fixed}.",
            )
            self._set_pending_bubble_number(fixed)
        finally:
            self._pending_number_adjusting = False

    def _ensure_notes_dialog(self) -> NotesExtractDialog:
        if self.notes_dialog is None:
            self.notes_dialog = NotesExtractDialog(self)
            try:
                self.notes_dialog.insert_to_form3_requested.connect(self._on_notes_insert_to_form3)
            except Exception:
                pass
        return self.notes_dialog

    def _on_notes_insert_to_form3(self, text: str, source_dialog=None) -> None:
        s = str(text or "").strip()
        if not s:
            return
        try:
            self.insert_notes_to_form3_requested.emit(s, source_dialog)
        except Exception:
            pass

    def set_note_region_mode(self, enabled: bool) -> None:
        if enabled:
            if getattr(self, "add_range_btn", None) is not None and self.add_range_btn.isChecked():
                self.add_range_btn.setChecked(False)
            # Disable bubble placement when selecting regions
            if self.placing_mode:
                self.placing_mode = False
                self._range_end_number = None
                self.add_bubble_btn.setChecked(False)
                self.add_bubble_btn.setText("Add Bubble")
                self.view.set_placing_mode(False)
        self.view.set_note_region_mode(enabled)
        self.add_note_region_btn.setChecked(enabled)

    def toggle_note_region_mode(self) -> None:
        self.set_note_region_mode(self.add_note_region_btn.isChecked())

    def on_note_region_created(self, rect: QRectF) -> None:
        if self.pixmap_item is None:
            return

        page_rect = self.pixmap_item.boundingRect()
        if page_rect.width() <= 0 or page_rect.height() <= 0:
            return

        x0 = max(0.0, min(1.0, rect.left() / page_rect.width()))
        y0 = max(0.0, min(1.0, rect.top() / page_rect.height()))
        x1 = max(0.0, min(1.0, rect.right() / page_rect.width()))
        y1 = max(0.0, min(1.0, rect.bottom() / page_rect.height()))
        if x1 <= x0 or y1 <= y0:
            return

        # Store regions in unrotated PDF-normalized coordinates.
        rot = self._current_page_rotation()
        if rot:
            inv = (-rot) % 360
            x0, y0, x1, y1 = self._rotate_norm_rect(x0, y0, x1, y1, inv)

        # Only allow one Notes Window per page: creating a new one clears the previous.
        self.note_regions_by_page[self.current_page] = [(x0, y0, x1, y1)]
        self._rebuild_note_region_items()

    def _update_note_region_from_item(self, index0: int, rect_scene: QRectF) -> None:
        """Update stored normalized (unrotated) coords for a moved/resized region item."""
        if self.pixmap_item is None:
            return
        try:
            page_rect = self.pixmap_item.boundingRect()
            if page_rect.width() <= 0 or page_rect.height() <= 0:
                return
        except Exception:
            return

        # Map scene rect back into pixmap-item coordinates.
        try:
            tl = self.pixmap_item.mapFromScene(rect_scene.topLeft())
            br = self.pixmap_item.mapFromScene(rect_scene.bottomRight())
            rect_view = QRectF(tl, br).normalized()
        except Exception:
            rect_view = rect_scene

        try:
            x0 = max(0.0, min(1.0, rect_view.left() / page_rect.width()))
            y0 = max(0.0, min(1.0, rect_view.top() / page_rect.height()))
            x1 = max(0.0, min(1.0, rect_view.right() / page_rect.width()))
            y1 = max(0.0, min(1.0, rect_view.bottom() / page_rect.height()))
        except Exception:
            return

        if x1 <= x0 or y1 <= y0:
            return

        rot = self._current_page_rotation()
        if rot:
            inv = (-rot) % 360
            try:
                x0, y0, x1, y1 = self._rotate_norm_rect(x0, y0, x1, y1, inv)
            except Exception:
                pass

        regions = list(self.note_regions_by_page.get(self.current_page, []))
        if not regions:
            regions = [(x0, y0, x1, y1)]
        else:
            i = int(index0)
            if i < 0:
                i = 0
            if i >= len(regions):
                i = len(regions) - 1
            regions[i] = (x0, y0, x1, y1)
        self.note_regions_by_page[self.current_page] = regions

    def _rebuild_note_region_items(self) -> None:
        # Clear existing items
        for item in self.note_region_items:
            try:
                self.scene.removeItem(item)
            except Exception:
                pass
        self.note_region_items = []

        if self.pixmap_item is None:
            return

        page_rect = self.pixmap_item.boundingRect()
        regions = self.note_regions_by_page.get(self.current_page, [])
        rot = self._current_page_rotation()
        for idx, (x0, y0, x1, y1) in enumerate(regions, start=1):
            # Convert stored unrotated coords -> current view coords
            if rot:
                x0, y0, x1, y1 = self._rotate_norm_rect(x0, y0, x1, y1, rot)
            r = QRectF(
                page_rect.width() * x0,
                page_rect.height() * y0,
                page_rect.width() * (x1 - x0),
                page_rect.height() * (y1 - y0),
            )
            item = _NoteRegionItem(r, viewer=self, index0=idx - 1)
            item.setPen(QPen(QColor(255, 200, 0), 2))
            item.setBrush(QBrush(QColor(255, 200, 0, 40)))
            # Store index for deletion
            item.setData(0, idx - 1)
            self.scene.addItem(item)
            self.note_region_items.append(item)

    def clear_note_regions(self) -> None:
        self.note_regions_by_page[self.current_page] = []
        self._rebuild_note_region_items()

    def clear_all_note_regions(self) -> None:
        try:
            self.note_regions_by_page = {}
        except Exception:
            pass
        self._rebuild_note_region_items()

    def clear_extracted_notes_dialog(self) -> None:
        try:
            if self.notes_dialog is not None:
                self.notes_dialog.clear_content()
        except Exception:
            pass

    def delete_selected_note_regions(self) -> None:
        if not self.note_region_items:
            return
        selected_indexes = sorted(
            {
                int(it.data(0))
                for it in self.note_region_items
                if it.isSelected() and it.data(0) is not None
            },
            reverse=True,
        )
        if not selected_indexes:
            return

        regions = self.note_regions_by_page.get(self.current_page, [])
        for i in selected_indexes:
            if 0 <= i < len(regions):
                regions.pop(i)
        self.note_regions_by_page[self.current_page] = regions
        self._rebuild_note_region_items()

    def on_bubble_click(self, scene_pos):
        """Handle click in placement mode."""
        if self.placing_mode:
            if self.range_mode:
                try:
                    self._set_pending_bubble_number(self._lowest_available_number())
                except Exception:
                    pass
                result = self._prompt_add_range(int(self._pending_bubble_number()))
                if result is None:
                    return

                start, end = result
                self.add_range_bubble(int(start), int(end), scene_pos.x(), scene_pos.y())

                try:
                    self._set_pending_bubble_number(self._lowest_available_number())
                except Exception:
                    pass

                # Stay in range mode so user can keep placing ranges.
                # (Next start is automatically updated from persisted bubbles.)
                self.placing_mode = True
                self.range_mode = True
                self._range_end_number = None
                try:
                    self.add_range_btn.setChecked(True)
                except Exception:
                    pass
                try:
                    self.view.set_placing_mode(True)
                except Exception:
                    pass
                return

            # Normal Add Bubble mode
            n = int(self._pending_bubble_number())
            overlap = self._range_overlap(n, n)
            if overlap:
                fixed = self._lowest_available_number()
                QMessageBox.warning(
                    self,
                    "Duplicate Bubble Number",
                    f"Bubble number {n} already exists. Next available is {fixed}.",
                )
                self._set_pending_bubble_number(fixed)
                return

            self.add_bubble(int(n), scene_pos.x(), scene_pos.y())
            try:
                self._set_pending_bubble_number(self._lowest_available_number())
            except Exception:
                pass
            self.add_bubble_btn.setText("Click to place")

    def load_pdf(self, file_path):
        """Load a PDF file."""
        try:
            opened_path = str(file_path)
            self._opened_pdf_path = opened_path
            self._sidecar_context_pdf_path = opened_path

            if getattr(self, "_debug_enabled", False):
                try:
                    print(f"[AS9102_DEBUG_PDF] PdfViewer.load_pdf(opened={opened_path})", flush=True)
                except Exception:
                    pass

            # Save target: only auto-set when reopening an annotated drawing (sidecar exists),
            # otherwise force Save to behave like Save As on first save.
            self._save_target_pdf_path = None

            # If we're opening a previously saved/annotated PDF that has a sidecar,
            # render using the original (clean) PDF as the background to avoid
            # seeing the baked-in bubbles + editable bubbles at the same time.
            sidecar_data = self._read_sidecar_json(opened_path) or None
            source_path = opened_path
            if isinstance(sidecar_data, dict):
                # If a sidecar exists for this opened PDF, treat it as the save target.
                try:
                    self._save_target_pdf_path = str(opened_path)
                except Exception:
                    self._save_target_pdf_path = None
                try:
                    candidate = str(sidecar_data.get("source_pdf_path") or "").strip()
                except Exception:
                    candidate = ""
                # DISABLE source path fallback.
                # When opening a saved PDF, we want to load THAT PDF (which contains the correct
                # embedded annotations and deletions), not the original source PDF (which might
                # still have the deleted "External" bubbles).
                # The viewer logic is smart enough to hide the embedded internal bubbles
                # and show the editable overlays.
                # if candidate and os.path.exists(candidate):
                #     source_path = candidate

            if getattr(self, "_debug_enabled", False):
                try:
                    has_sidecar = isinstance(sidecar_data, dict)
                    print(f"[AS9102_DEBUG_PDF]  source_path={source_path} sidecar={has_sidecar}", flush=True)
                except Exception:
                    pass

            self.file_path = source_path
            self.doc = fitz.open(source_path)
            self.total_pages = len(self.doc)
            self.current_page = 0

            if getattr(self, "_debug_enabled", False):
                try:
                    print(f"[AS9102_DEBUG_PDF]  pages={self.total_pages}", flush=True)
                except Exception:
                    pass
            
            self.page_combo.blockSignals(True)
            self.page_combo.clear()
            for i in range(self.total_pages):
                self.page_combo.addItem(f"{i + 1} / {self.total_pages}")
            self.page_combo.blockSignals(False)
            
            self._did_initial_fit = False
            self.note_regions_by_page = {}
            self.page_rotation_by_page = {}
            self._last_render_rotation = 0
            self._rendered_page_index = None

            # Clear any previously rendered page/scene so overlap/persist logic
            # can't accidentally write empty bubble specs over the new import.
            try:
                self.pixmap_item = None
            except Exception:
                pass
            try:
                self.page_items = []
            except Exception:
                pass
            try:
                if getattr(self, "scene", None) is not None:
                    self.scene.clear()
            except Exception:
                pass

            self.bubble_specs_by_page = {}
            self.bubbles = []
            self.next_bubble_number = 1
            self.placing_mode = False
            self._range_end_number = None
            self.range_mode = False
            self._undo_stack = []

            # Load editable bubble state if it exists for the OPENED PDF.
            self._load_edit_state_sidecar(opened_path, data=sidecar_data)

            # Always try to import bubbles from embedded PDF annotations and overlay them.
            # Important: if we rendered from a clean "source" PDF, the embedded annotations
            # (Kofax, etc.) may exist only in the OPENED PDF. Import from the opened PDF.
            try:
                import_doc = None
                try:
                    if str(opened_path) and os.path.exists(opened_path) and str(opened_path) != str(source_path):
                        import_doc = fitz.open(opened_path)
                except Exception:
                    import_doc = None
                if import_doc is None:
                    import_doc = self.doc

                imported = self._extract_bubbles_from_doc_annotations(import_doc)
                if getattr(self, "_debug_enabled", False):
                    try:
                        page_cnt = len(imported or {})
                        spec_cnt = sum(len(v or []) for v in (imported or {}).values())
                        print(f"[AS9102_DEBUG_PDF]  imported_pages={page_cnt} imported_specs={spec_cnt}", flush=True)
                    except Exception:
                        pass
                if imported:
                    # If no sidecar exists, treat imported embedded annotations as the initial bubble set.
                    # (This avoids edge cases where overlap logic can accidentally skip everything on reload.)
                    try:
                        has_sidecar = isinstance(sidecar_data, dict)
                    except Exception:
                        has_sidecar = False

                    try:
                        existing_specs_count = sum(len(v or []) for v in (self.bubble_specs_by_page or {}).values())
                    except Exception:
                        existing_specs_count = 0

                    if (not has_sidecar) and existing_specs_count == 0:
                        imported_spec_cnt = 0
                        try:
                            imported_spec_cnt = sum(len(v or []) for v in (imported or {}).values())
                        except Exception:
                            imported_spec_cnt = 0

                        assign_err = None
                        try:
                            new_specs: dict[int, list[tuple[int, int, float, float, int, str]]] = {}
                            for k, v in (imported or {}).items():
                                try:
                                    page_k = int(k)
                                except Exception:
                                    continue
                                try:
                                    new_specs[page_k] = list(v or [])
                                except Exception:
                                    new_specs[page_k] = []
                            self.bubble_specs_by_page = new_specs
                        except Exception as e:
                            assign_err = e
                            # Fall back to direct assignment.
                            try:
                                self.bubble_specs_by_page = dict(imported or {})
                            except Exception:
                                pass

                        # If something went wrong and we ended up with no specs, fall back.
                        try:
                            assigned_specs_cnt = sum(len(v or []) for v in (self.bubble_specs_by_page or {}).values())
                        except Exception:
                            assigned_specs_cnt = 0
                        if imported_spec_cnt > 0 and assigned_specs_cnt == 0:
                            try:
                                self.bubble_specs_by_page = dict(imported or {})
                            except Exception:
                                pass
                        try:
                            self._recompute_next_bubble_number()
                        except Exception:
                            pass
                        if getattr(self, "_debug_enabled", False):
                            try:
                                final_specs = sum(len(v or []) for v in (self.bubble_specs_by_page or {}).values())
                                pages = list((self.bubble_specs_by_page or {}).keys())
                                if assign_err is not None:
                                    print(f"[AS9102_DEBUG_PDF]  import_mode=replace assign_error={assign_err}", flush=True)
                                print(f"[AS9102_DEBUG_PDF]  import_mode=replace pages={pages} final_specs={final_specs}", flush=True)
                            except Exception:
                                pass
                    else:
                        # Merge without overwriting existing bubbles; skip duplicates.
                        used = set(self.get_bubbled_numbers() or set())
                        added = 0
                        skipped_overlap = 0
                        changed = False
                        for page_index, specs in (imported or {}).items():
                            if not specs:
                                continue
                            cur = list(self.bubble_specs_by_page.get(int(page_index), []) or [])
                            for spec in specs:
                                try:
                                    start, end, rx, ry, br = spec[:5]
                                    bf = spec[5] if len(spec) > 5 else ""
                                except Exception:
                                    continue
                                try:
                                    s = int(start)
                                    e = int(end)
                                except Exception:
                                    continue
                                if s <= 0:
                                    continue
                                if e < s:
                                    e = s
                                # Skip if any number overlaps.
                                overlap = any((n in used) for n in range(s, e + 1))
                                if overlap:
                                    skipped_overlap += 1
                                    continue
                                cur.append((int(s), int(e), float(rx), float(ry), int(br), str(bf or "")))
                                for n in range(s, e + 1):
                                    used.add(int(n))
                                added += 1
                                changed = True
                            if cur:
                                cur.sort(key=lambda t: (t[0], t[1], t[2], t[3]))
                                self.bubble_specs_by_page[int(page_index)] = cur

                        # If merge produced nothing but import found specs, fall back to replacing.
                        try:
                            after_specs = sum(len(v or []) for v in (self.bubble_specs_by_page or {}).values())
                        except Exception:
                            after_specs = 0
                        if after_specs == 0:
                            try:
                                self.bubble_specs_by_page = {
                                    int(k): list(v or [])
                                    for k, v in (imported or {}).items()
                                }
                                changed = True
                            except Exception:
                                pass

                        if changed:
                            self._recompute_next_bubble_number()

                        if getattr(self, "_debug_enabled", False):
                            try:
                                print(f"[AS9102_DEBUG_PDF]  import_mode=merge added={added} skipped_overlap={skipped_overlap}", flush=True)
                            except Exception:
                                pass

                if getattr(self, "_debug_enabled", False):
                    try:
                        final_specs = sum(len(v or []) for v in (self.bubble_specs_by_page or {}).values())
                        final_nums = len(self.get_bubbled_numbers() or set())
                        print(f"[AS9102_DEBUG_PDF]  final_specs={final_specs} final_numbers={final_nums}", flush=True)
                    except Exception:
                        pass
            except Exception:
                pass
            finally:
                try:
                    if 'import_doc' in locals() and import_doc is not None and import_doc is not self.doc:
                        import_doc.close()
                except Exception:
                    pass
            self._set_dirty(False)
            self._last_saved_pdf_path = str(opened_path)

            # Initial render at baseline (100%) then fit once the widget has size
            self.set_zoom(1.0, render_now=True)
            QTimer.singleShot(50, self.fit_to_view)

            # Notify listeners that bubbles for this drawing are now known.
            try:
                self.bubbles_changed.emit(self.get_bubbled_numbers())
            except Exception:
                pass
            
        except Exception as e:
            logger.exception("Error loading PDF")
            QMessageBox.critical(self, "Error", f"Failed to load PDF:\n{e}")

    def get_annotation_counts(self, page_index: int) -> tuple[int, int]:
        """Return (internal_count, external_count) for the given page.
        
        internal_count: Number of active/editable bubbles in the app for this page.
        external_count: Number of embedded annotations in the PDF that are NOT internal.
        """
        # Internal: Count the active bubbles in memory
        internal = 0
        try:
            specs = self.bubble_specs_by_page.get(int(page_index), [])
            internal = len(specs)
        except Exception:
            pass

        # External: Count embedded annotations that are not tagged as internal
        external = 0
        if self.doc:
            try:
                page = self.doc.load_page(int(page_index))
                for ann in (page.annots() or []):
                    try:
                        info = getattr(ann, "info", {}) or {}
                        # If it's tagged internal, we ignore it for the "External" count
                        # (It's either a duplicate of an active bubble, or hidden)
                        if info.get("title") == "AS9102_FAI_BUBBLE" or info.get("subject") == "AS9102_FAI_BUBBLE":
                            continue
                        external += 1
                    except Exception:
                        pass
            except Exception:
                pass
        
        return (internal, external)

    def get_total_annotation_counts(self) -> tuple[int, int]:
        """Return (total_internal, total_external) across all pages."""
        # Total Internal: Sum of all active bubbles in memory
        total_internal = 0
        try:
            for specs in self.bubble_specs_by_page.values():
                total_internal += len(specs)
        except Exception:
            pass
        
        # Total External: Sum of all non-internal embedded annotations
        total_external = 0
        if self.doc:
            try:
                for page_index in range(self.doc.page_count):
                    try:
                        page = self.doc.load_page(page_index)
                        for ann in (page.annots() or []):
                            info = getattr(ann, "info", {}) or {}
                            if info.get("title") == "AS9102_FAI_BUBBLE" or info.get("subject") == "AS9102_FAI_BUBBLE":
                                continue
                            total_external += 1
                    except Exception:
                        pass
            except Exception:
                pass
            
        return (total_internal, total_external)

    def toggle_enhance_mode(self):
        try:
            enabled = bool(self.enhance_btn.isChecked())
        except Exception:
            enabled = False
        self.set_enhance_mode(bool(enabled))

    def set_enhance_mode(self, enabled: bool) -> None:
        """Enable/disable image enhancement without requiring the Enhance button."""
        self.enhance_mode = bool(enabled)
        try:
            self.enhance_panel.setVisible(bool(self.enhance_mode))
        except Exception:
            pass
        try:
            if getattr(self, "enhance_btn", None) is not None:
                self.enhance_btn.blockSignals(True)
                self.enhance_btn.setChecked(bool(self.enhance_mode))
                self.enhance_btn.blockSignals(False)
        except Exception:
            try:
                self.enhance_btn.blockSignals(False)
            except Exception:
                pass
        try:
            self.render_current_page(preserve_center=True)
        except Exception:
            pass

    def on_enhancement_changed(self):
        self.brightness = self.brightness_slider.value() / 10.0
        self.contrast = self.contrast_slider.value() / 10.0
        self.sharpness = self.sharpness_slider.value() / 10.0
        self.render_current_page(preserve_center=True)

    def reset_enhancements(self):
        self.brightness_slider.setValue(10)
        self.contrast_slider.setValue(10)
        self.sharpness_slider.setValue(10)
        self.on_enhancement_changed()

    def render_current_page(self, target_scale=2.0, preserve_center: bool = True, source_rotation: int | None = None):
        """Render the current page at the given scale factor.
        
        target_scale: multiplier for base DPI (2.0 = 144 DPI for good quality)
        """
        if not self.doc:
            return

        target_rotation = self._current_page_rotation()
        if source_rotation is None:
            source_rotation = target_rotation
        source_rotation = int(source_rotation) % 360
        delta_rotation = (target_rotation - source_rotation) % 360

        # Persist bubbles for the current page only when the scene currently
        # corresponds to the same page (avoids writing old-page bubbles into
        # the new page during page switches).
        if self._rendered_page_index == self.current_page:
            self._persist_current_page_bubbles(rotation=source_rotation)
        
        # Preserve the current view center as a relative position on the page
        center_rx = 0.5
        center_ry = 0.5
        if preserve_center and self.pixmap_item is not None:
            page_rect = self.pixmap_item.boundingRect()
            if page_rect.width() > 0 and page_rect.height() > 0:
                center_scene = self.view.mapToScene(self.view.viewport().rect().center())
                center_rx = float(center_scene.x() / page_rect.width())
                center_ry = float(center_scene.y() / page_rect.height())
                if delta_rotation:
                    center_rx, center_ry = self._rotate_norm_point(center_rx, center_ry, delta_rotation)

        self.scene.clear()
        self.bubbles = []
        self.page_items = []
        
        page = self.doc.load_page(self.current_page)
        
        # Render at scaled DPI for quality
        # target_scale of 2.0 gives us 144 effective DPI which is good for most screens
        render_dpi = self.base_dpi * target_scale
        zoom = render_dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        
        # --- Visibility Logic ---
        # Temporarily modify annotation flags to control visibility during render.
        # We restore them immediately after.
        original_flags = {}
        try:
            annots = list(page.annots() or [])
            for ann in annots:
                try:
                    original_flags[ann.xref] = ann.flags
                    info = getattr(ann, "info", {}) or {}
                    is_internal = (info.get("title") == "AS9102_FAI_BUBBLE" or info.get("subject") == "AS9102_FAI_BUBBLE")
                    
                    should_show = False
                    if is_internal:
                        should_show = self.show_internal_annots
                    else:
                        should_show = self.show_external_annots
                    
                    # To show: clear HIDDEN/INVISIBLE. To hide: set HIDDEN.
                    # CRITICAL FIX: If we are about to draw a widget for this bubble (because it's in our specs),
                    # we MUST hide the underlying PDF annotation regardless of 'should_show'.
                    # Otherwise we get "double vision" (burned-in image + widget).
                    
                    if is_internal:
                        # Always hide internal annotations to prevent ghosts and double-vision.
                        # The Widget layer is the source of truth for internal bubbles.
                        # If the user wants to see them, we show the widgets.
                        # If the user wants to hide them, we hide the widgets.
                        # We NEVER want to see the baked-in annotation for an internal bubble during an active session.
                        new_flags = ann.flags | fitz.PDF_ANNOT_IS_HIDDEN
                    elif should_show:
                        new_flags = ann.flags & ~fitz.PDF_ANNOT_IS_HIDDEN & ~fitz.PDF_ANNOT_IS_INVISIBLE
                    else:
                        new_flags = ann.flags | fitz.PDF_ANNOT_IS_HIDDEN
                    
                    if new_flags != ann.flags:
                        ann.set_flags(new_flags)
                        ann.update() # Commit for render
                except Exception:
                    pass
        except Exception:
            pass
        # ------------------------

        # Some PyMuPDF builds don't support rotate= keyword.
        # Render unrotated then rotate the pixmap via Qt.
        # Render optionally with the PDF's own annotations so third-party markup (e.g. Kofax) can be shown/hidden.
        try:
            # Always pass annots=True because we are controlling visibility via flags now.
            pix = page.get_pixmap(matrix=mat, alpha=False, annots=True)
        except TypeError:
            # Older PyMuPDF versions may not support the annots= keyword.
            # Fall back to defaults.
            pix = page.get_pixmap(matrix=mat, alpha=False)
        
        # --- Restore Flags ---
        try:
            for ann in annots:
                try:
                    if ann.xref in original_flags:
                        if ann.flags != original_flags[ann.xref]:
                            ann.set_flags(original_flags[ann.xref])
                            ann.update()
                except Exception:
                    pass
        except Exception:
            pass
        # ---------------------

        # Create QImage and QPixmap
        do_enhance = False
        try:
            do_enhance = bool(self.enhance_mode) and (
                float(getattr(self, "brightness", 1.0) or 1.0) != 1.0
                or float(getattr(self, "contrast", 1.0) or 1.0) != 1.0
                or float(getattr(self, "sharpness", 1.0) or 1.0) != 1.0
            )
        except Exception:
            do_enhance = bool(self.enhance_mode)

        if do_enhance:
             # Convert to PIL
             mode = "RGB" if pix.n == 3 else "RGBA"
             # pix.samples is bytes
             pil_img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
             
             # Apply enhancements
             if self.brightness != 1.0:
                 pil_img = ImageEnhance.Brightness(pil_img).enhance(self.brightness)
             if self.contrast != 1.0:
                 pil_img = ImageEnhance.Contrast(pil_img).enhance(self.contrast)
             if self.sharpness != 1.0:
                 pil_img = ImageEnhance.Sharpness(pil_img).enhance(self.sharpness)
                 
             # Convert back to QImage
             # QImage needs data to remain valid, so we keep a reference if needed, 
             # but here we create QPixmap immediately so it should be fine.
             # However, QImage(bytes, ...) does not copy data, so we must ensure 'data' stays alive
             # until QPixmap is created.
             data = pil_img.tobytes("raw", "RGB")
             img = QImage(data, pil_img.width, pil_img.height, QImage.Format_RGB888)
             # We must copy the image to ensure it owns the data, because 'data' variable will go out of scope
             img = img.copy() 
        else:
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)

        pixmap = QPixmap.fromImage(img)

        if target_rotation:
            t = QTransform()
            t.rotate(float(target_rotation))
            pixmap = pixmap.transformed(t, Qt.SmoothTransformation)
        
        self.pixmap_item = self.scene.addPixmap(pixmap)
        self.pixmap_item.setZValue(0)
        # Store the scale used for this render
        self.current_render_scale = target_scale
        self.page_items.append(self.pixmap_item)
        
        self.scene.setSceneRect(self.pixmap_item.boundingRect())

        self._last_render_rotation = target_rotation
        self._rendered_page_index = self.current_page

        # Recreate bubbles for THIS page only (stored in unrotated normalized coords)
        page_rect = self.pixmap_item.boundingRect()
        if page_rect.width() > 0 and page_rect.height() > 0:
            specs = self.bubble_specs_by_page.get(self.current_page, [])
            for spec in specs:
                try:
                    start, end, rx0, ry0, br = spec[:5]
                    bf = spec[5] if len(spec) > 5 else ""
                except Exception:
                    continue
                rx, ry = float(rx0), float(ry0)
                if target_rotation:
                    rx, ry = self._rotate_norm_point(rx, ry, target_rotation)
                x = rx * page_rect.width()
                y = ry * page_rect.height()
                label = f"{int(start)}-{int(end)}" if int(end) > int(start) else str(int(start))
                b = BubbleItem(int(start), x, y, base_radius=int(br), parent_viewer=self, range_end=int(end), display_text=label, backfill_rgb=str(bf or ""))
                
                # Also respect internal visibility for the interactive overlay
                b.setVisible(self.show_internal_annots)
                
                self.scene.addItem(b)
                self.bubbles.append(b)

        # Rebuild note region overlays for this page
        self._rebuild_note_region_items()

        # Rebuild drawing grid overlay (if enabled)
        self._rebuild_grid_overlay()

        # Restore center
        page_rect = self.pixmap_item.boundingRect()
        self.view.centerOn(page_rect.width() * center_rx, page_rect.height() * center_ry)
        
        # Emit signal to update UI counts
        try:
            self.bubbles_changed.emit(self.get_bubbled_numbers())
        except Exception:
            pass

    def go_to_page(self, index):
        if 0 <= index < self.total_pages:
            if index != self.current_page:
                self._persist_current_page_bubbles(rotation=self._last_render_rotation)
            self.current_page = index
            self.render_current_page(target_scale=self.base_render_scale * self._zoom_factor, source_rotation=self._last_render_rotation)

    def prev_page(self):
        if not self.doc:
            return
        if self.current_page > 0:
            self._persist_current_page_bubbles(rotation=self._last_render_rotation)
            self.current_page -= 1
            self.page_combo.setCurrentIndex(self.current_page)
            self.render_current_page(target_scale=self.current_render_scale)

    def next_page(self):
        if not self.doc:
            return
        if self.current_page < self.total_pages - 1:
            self._persist_current_page_bubbles(rotation=self._last_render_rotation)
            self.current_page += 1
            self.page_combo.setCurrentIndex(self.current_page)
            self.render_current_page(target_scale=self.current_render_scale)

    def zoom_in(self):
        self.set_zoom(self._zoom_factor * 1.25)

    def zoom_out(self):
        self.set_zoom(self._zoom_factor / 1.25)

    def zoom_100(self):
        self.set_zoom(1.0)

    def get_zoom_factor(self) -> float:
        return float(self._zoom_factor)

    def set_zoom(self, zoom_factor: float, render_now: bool = True) -> None:
        try:
            zoom_factor = float(zoom_factor)
        except Exception:
            return
        zoom_factor = max(self.MIN_ZOOM, min(zoom_factor, self.MAX_ZOOM))
        if abs(zoom_factor - self._zoom_factor) < 1e-6 and self.pixmap_item is not None:
            self._update_zoom_label()
            return

        self._zoom_factor = zoom_factor
        self._update_zoom_label()
        if render_now:
            self.render_current_page(target_scale=self.base_render_scale * self._zoom_factor)

    def set_show_pdf_annots(self, enabled: bool) -> None:
        self.show_pdf_annots = bool(enabled)
        try:
            self._settings.setValue("pdf_viewer/show_pdf_annots", bool(self.show_pdf_annots))
        except Exception:
            pass
        try:
            self._rendered_page_index = None
        except Exception:
            pass
        self.render_current_page(target_scale=self.current_render_scale, source_rotation=self._last_render_rotation)

    def set_bubble_color(self, color: QColor) -> None:
        if color is None:
            return
        try:
            if not color.isValid():
                return
        except Exception:
            pass
        self.bubble_color = QColor(color)
        try:
            self._settings.setValue("pdf_viewer/bubble_color", str(self.bubble_color.name()[1:]))
        except Exception:
            pass
        for b in list(getattr(self, "bubbles", []) or []):
            try:
                b.update()
            except Exception:
                pass

    def set_bubble_backfill_white(self, enabled: bool) -> None:
        """Toggle white-filled bubbles (backfill) for readability on drawings."""
        try:
            self.bubble_backfill_white = bool(enabled)
        except Exception:
            self.bubble_backfill_white = True
        try:
            self._settings.setValue("pdf_viewer/bubble_backfill_white", bool(self.bubble_backfill_white))
        except Exception:
            pass
        try:
            self._settings.setValue("pdf_viewer/bubble_backfill_rgb", str(getattr(self, "bubble_backfill_rgb", "FFFFFF")))
        except Exception:
            pass
        for b in list(getattr(self, "bubbles", []) or []):
            try:
                b.update()
            except Exception:
                pass

    def _sanitize_rgb_hex(self, rgb: str | None, default: str = "FFFFFF") -> str:
        s = str(rgb or "").strip()
        s = re.sub(r"[^0-9a-fA-F]", "", s)
        if len(s) != 6:
            s = str(default or "FFFFFF")
            s = re.sub(r"[^0-9a-fA-F]", "", s)
            if len(s) != 6:
                s = "FFFFFF"
        return s.upper()

    def get_bubble_backfill_qcolor(self):
        """Return QColor for bubble backfill if enabled, else None."""
        try:
            if not bool(getattr(self, "bubble_backfill_white", False)):
                return None
        except Exception:
            return None
        try:
            rgb = self._sanitize_rgb_hex(getattr(self, "bubble_backfill_rgb", "FFFFFF"), "FFFFFF")
        except Exception:
            rgb = "FFFFFF"
        try:
            qc = QColor("#" + str(rgb))
            if qc.isValid():
                return qc
        except Exception:
            pass
        return QColor("#FFFFFF")

    def set_bubble_backfill_color(self, rgb: str | None, *, enabled: bool | None = None) -> None:
        """Set bubble backfill color (hex RGB) and optionally enable/disable."""
        try:
            rgb = self._sanitize_rgb_hex(rgb, "FFFFFF")
        except Exception:
            rgb = "FFFFFF"

        try:
            if enabled is None:
                enabled = True
        except Exception:
            enabled = True

        try:
            self.bubble_backfill_rgb = str(rgb)
        except Exception:
            self.bubble_backfill_rgb = "FFFFFF"

        try:
            if enabled is not None:
                self.bubble_backfill_white = bool(enabled)
        except Exception:
            pass

        try:
            self._settings.setValue("pdf_viewer/bubble_backfill_rgb", str(self.bubble_backfill_rgb))
            self._settings.setValue("pdf_viewer/bubble_backfill_white", bool(self.bubble_backfill_white))
        except Exception:
            pass

        for b in list(getattr(self, "bubbles", []) or []):
            try:
                b.update()
            except Exception:
                pass

    def apply_backfill_to_selected_bubbles(self, rgb: str | None) -> bool:
        """Apply backfill color to selected bubbles only."""
        try:
            selected = [b for b in (self.bubbles or []) if hasattr(b, "isSelected") and b.isSelected()]
        except Exception:
            selected = []
        if not selected:
            return False

        try:
            rgb_norm = self._sanitize_rgb_hex(rgb, "FFFFFF") if rgb else ""
        except Exception:
            rgb_norm = ""

        try:
            self._push_undo_state()
        except Exception:
            pass

        for b in selected:
            try:
                b.backfill_rgb = str(rgb_norm or "")
                b.update()
            except Exception:
                pass

        try:
            self._persist_current_page_bubbles()
        except Exception:
            pass
        try:
            self._set_dirty(True)
        except Exception:
            pass
        return True

    def delete_pdf_annotations(self, *, all_pages: bool = False) -> int:
        """Delete existing PDF annotations from the loaded document.

        Returns count deleted (best-effort).
        """
        if not self.doc:
            return 0
        deleted = 0
        try:
            page_indexes = range(int(self.total_pages or 0)) if all_pages else [int(self.current_page)]
        except Exception:
            page_indexes = [int(self.current_page)]

        for page_index in page_indexes:
            try:
                page = self.doc.load_page(int(page_index))
            except Exception:
                continue

            try:
                annots = list(page.annots() or [])
            except Exception:
                annots = []

            for ann in annots:
                try:
                    # Prefer page.delete_annot(ann) if available.
                    if hasattr(page, "delete_annot"):
                        page.delete_annot(ann)
                    elif hasattr(ann, "delete"):
                        ann.delete()
                    else:
                        continue
                    deleted += 1
                except Exception:
                    continue

        if deleted:
            self._pdf_annots_mutated = True
            if all_pages:
                self._pdf_annots_deleted_all = True
                try:
                    self._pdf_annots_deleted_pages = set()
                except Exception:
                    pass
            else:
                try:
                    self._pdf_annots_deleted_pages.add(int(self.current_page))
                except Exception:
                    pass
            self._set_dirty(True)
            try:
                # Cache mutated bytes so Save/Save As can reliably use the in-memory state.
                # Use a full-save snapshot (garbage collected) so deleted annots can't reappear.
                if self._debug_enabled:
                    print(f"DEBUG: Snapshotting PDF bytes after deletion (garbage=0)...")
                
                # Force a reload from bytes to ensure the in-memory doc is clean
                if self.doc is not None:
                    # 1. Snapshot current state
                    b = self._snapshot_pdf_bytes(self.doc)
                    if b:
                        # 2. Reload self.doc from these bytes so it's fresh
                        try:
                            new_doc = fitz.open(stream=b, filetype="pdf")
                            self.doc.close()
                            self.doc = new_doc
                            if self._debug_enabled:
                                print("DEBUG: Reloaded self.doc from snapshot bytes.")
                        except Exception as e:
                            if self._debug_enabled:
                                print(f"DEBUG: Failed to reload self.doc: {e}")
                        
                        self._pdf_doc_bytes_after_annots_mutation = b
                        if self._debug_enabled:
                            print(f"DEBUG: Snapshot success, bytes len: {len(b)}")
                    else:
                        self._pdf_doc_bytes_after_annots_mutation = None
            except Exception as e:
                if self._debug_enabled:
                    print(f"DEBUG: Snapshot failed: {e}")
                self._pdf_doc_bytes_after_annots_mutation = None
            try:
                self._rendered_page_index = None
            except Exception:
                pass
            self.render_current_page(target_scale=self.current_render_scale, source_rotation=self._last_render_rotation)
        return int(deleted)

    def _snapshot_pdf_bytes(self, doc: "fitz.Document") -> bytes | None:
        """Return a stable PDF byte snapshot of a document.

        We prefer a full save with garbage collection so deleted objects (like
        annotations) don't come back when the PDF is later saved/reopened.
        """
        if doc is None:
            return None
        try:
            import io

            buf = io.BytesIO()
            # Full rewrite with garbage collection to drop deleted objects.
            # garbage=0 is safest to avoid xref errors; garbage=4 was causing corruption.
            doc.save(buf, garbage=0, deflate=True)
            return buf.getvalue()
        except Exception:
            try:
                return doc.tobytes()
            except Exception:
                return None

    def _apply_deleted_pdf_annotations_to_doc(self, doc: "fitz.Document") -> None:
        """Apply the same embedded-annotation deletions to another doc instance.

        Save/Save As may re-open a file from disk; this ensures the user's deletions
        are enforced in the output even if we couldn't clone the mutated in-memory doc.
        """
        if doc is None:
            return

        try:
            if bool(getattr(self, "_pdf_annots_deleted_all", False)):
                page_indexes = range(int(getattr(doc, "page_count", 0) or 0))
            else:
                page_indexes = sorted(int(i) for i in (getattr(self, "_pdf_annots_deleted_pages", set()) or set()))
        except Exception:
            return

        for page_index in page_indexes:
            try:
                page = doc.load_page(int(page_index))
            except Exception:
                continue
            try:
                annots = list(page.annots() or [])
            except Exception:
                annots = []
            deleted_count = 0
            for ann in annots:
                try:
                    if hasattr(page, "delete_annot"):
                        page.delete_annot(ann)
                    elif hasattr(ann, "delete"):
                        ann.delete()
                    else:
                        continue
                    deleted_count += 1
                except Exception:
                    continue
            if self._debug_enabled and deleted_count > 0:
                print(f"DEBUG: Deleted {deleted_count} annotations from page {page_index} in output doc.")
    
    def _update_zoom_label(self):
        """Update the zoom percentage label."""
        percent = int(round(self._zoom_factor * 100))
        self.zoom_label.setText(f"{percent}%")

    def fit_to_view(self):
        """Fit the PDF page to the view."""
        if not self.doc:
            return

        # Compute a zoom factor that renders the page to fit the viewport.
        # With base_dpi=72, the pixmap size in pixels is: points * (base_render_scale * zoom_factor)
        page = self.doc.load_page(self.current_page)
        page_rect = page.rect
        if page_rect.width <= 0 or page_rect.height <= 0:
            return

        pw = float(page_rect.width)
        ph = float(page_rect.height)
        rot = self._current_page_rotation()
        if rot in (90, 270):
            pw, ph = ph, pw

        vp = self.view.viewport().size()
        vw = max(1, vp.width())
        vh = max(1, vp.height())
        fit_zoom = min(
            vw / (pw * self.base_render_scale),
            vh / (ph * self.base_render_scale),
        )
        self.set_zoom(fit_zoom)
        self._did_initial_fit = True

    def _pdf_text_in_clip(self, page: fitz.Page, clip: fitz.Rect | None) -> str:
        # Prefer block-based extraction and sort blocks top-to-bottom, left-to-right.
        # This improves ordering for engineering drawings where PDF internal reading
        # order is often non-visual (e.g., footers can appear first).
        try:
            blocks = page.get_text("blocks", clip=clip)
            parts: list[str] = []
            for b in sorted(blocks or [], key=lambda t: (round(float(t[1]), 2), round(float(t[0]), 2))):
                try:
                    txt = str(b[4] or "")
                except Exception:
                    txt = ""
                txt = txt.strip()
                if txt:
                    parts.append(txt)
            s = "\n".join(parts).strip()
            if s:
                return s
        except Exception:
            pass

        try:
            if clip is None:
                return page.get_text("text")
            return page.get_text("text", clip=clip)
        except Exception:
            return ""

    def _ocr_text_in_clip(self, page: fitz.Page, clip: fitz.Rect | None) -> str:
        try:
            import pytesseract
            from PIL import Image, ImageEnhance
            import io
        except Exception:
            QMessageBox.warning(
                self,
                "OCR Not Available",
                "OCR requires Pillow + pytesseract. Install packages and ensure Tesseract OCR is installed on Windows.",
            )
            return ""

        # Ensure pytesseract knows where the Tesseract binary is.
        # You can also override with env var `AS9102_TESSERACT_CMD`.
        try:
            from as9102_fai.ocr_utils import configure_pytesseract

            if not configure_pytesseract():
                QMessageBox.warning(
                    self,
                    "Tesseract Not Found",
                    "Tesseract OCR was not found. Install it or set AS9102_TESSERACT_CMD (e.g. C:\\Program Files\\Tesseract-OCR\\tesseract.exe).",
                )
                return ""
        except Exception:
            # If helper import fails, continue; pytesseract may still work via PATH.
            pass

        try:
            # Higher scale improves OCR for fine-print drawings.
            # We use 3.0 (approx 216 DPI) for better recognition of small text.
            ocr_scale = 3.0
            mat = fitz.Matrix(ocr_scale, ocr_scale)
            pix = page.get_pixmap(matrix=mat, clip=clip, alpha=False)
            
            # Convert directly to PIL Image (faster than png bytes)
            mode = "RGB" if pix.n == 3 else "RGBA"
            image = Image.frombytes(mode, [pix.width, pix.height], pix.samples)

            # Apply enhancements if enabled in UI
            if getattr(self, "enhance_mode", False):
                b = getattr(self, "brightness", 1.0)
                c = getattr(self, "contrast", 1.0)
                s = getattr(self, "sharpness", 1.0)
                
                if b != 1.0:
                    image = ImageEnhance.Brightness(image).enhance(b)
                if c != 1.0:
                    image = ImageEnhance.Contrast(image).enhance(c)
                if s != 1.0:
                    image = ImageEnhance.Sharpness(image).enhance(s)

            return pytesseract.image_to_string(image)
        except Exception as e:
            QMessageBox.warning(self, "OCR Failed", f"OCR failed: {e}")
            return ""

    def _normalize_notes_text(self, text: str) -> str:
        """Collapse whitespace, then add line breaks only for numbered notes (e.g. '1.' '2.')."""
        if not text:
            return ""

        # Collapse all whitespace (including newlines/tabs) into single spaces.
        s = re.sub(r"\s+", " ", str(text)).strip()
        if not s:
            return ""

        # Insert a newline before a note marker like " 12. " (but not decimals like 1.25).
        # We require a trailing space after the dot to qualify as a note marker.
        s = re.sub(r"(?<!^)\s+(?=\d{1,3}\.\s)", "\n", s)
        return s

    def extract_notes(self) -> None:
        if not self.doc:
            return

        mode = self.notes_mode_combo.currentText()

        # Only current page extraction (per UX request)
        pages = [self.current_page]

        chunks: list[str] = []
        used_sources: set[str] = set()
        for page_index in pages:
            page = self.doc.load_page(page_index)
            regions = self.note_regions_by_page.get(page_index, [])

            if not regions:
                QMessageBox.warning(
                    self,
                    "No Regions",
                    "Please click 'Notes Window' and draw at least one box around the notes before extracting.",
                )
                return

            page_rect = page.rect
            for region_index, (x0, y0, x1, y1) in enumerate(regions, start=1):
                clip = fitz.Rect(
                    page_rect.width * x0,
                    page_rect.height * y0,
                    page_rect.width * x1,
                    page_rect.height * y1,
                )
                text = ""

                if mode in ("Auto", "PDF Text"):
                    pdf_text = self._pdf_text_in_clip(page, clip) or ""
                    if mode == "PDF Text":
                        used_sources.add("PDF Text")
                        text = pdf_text
                    else:
                        if str(pdf_text).strip():
                            used_sources.add("PDF Text")
                            text = pdf_text

                if (mode in ("Auto", "OCR") and not str(text).strip()) or mode == "OCR":
                    # If user forced OCR, or Auto/PDF returned empty, fall back to OCR.
                    used_sources.add("OCR")
                    text = self._ocr_text_in_clip(page, clip)
                cleaned = self._normalize_notes_text(text)
                if cleaned:
                    chunks.append(f"Page {page_index + 1} / Region {region_index}:\n{cleaned}")

        # Separate regions clearly; keep numbered-note line breaks from normalization.
        content = "\n\n".join(chunks).strip() or "(No text found)"
        dlg = self._ensure_notes_dialog()

        if used_sources == {"OCR"}:
            src = "OCR"
        elif used_sources == {"PDF Text"}:
            src = "PDF Text"
        elif used_sources:
            src = "Mixed"
        else:
            src = ""

        dlg.set_source(src)
        dlg.append_content(content)
        dlg.show()
        dlg.raise_()
        dlg.activateWindow()

    def _persist_current_page_bubbles(self, rotation: int | None = None, *, recompute_next: bool = True) -> None:
        # Only persist if the scene corresponds to the current page.
        # This avoids writing empty specs during loads/re-renders.
        try:
            if getattr(self, "_rendered_page_index", None) != getattr(self, "current_page", None):
                return
        except Exception:
            pass
        if self.pixmap_item is None:
            return
        page_rect = self.pixmap_item.boundingRect()
        if page_rect.width() <= 0 or page_rect.height() <= 0:
            return

        rot = self._current_page_rotation() if rotation is None else (int(rotation) % 360)
        specs: list[tuple[int, int, float, float, int, str]] = []
        for b in self.bubbles:
            rx = float(b.pos().x() / page_rect.width())
            ry = float(b.pos().y() / page_rect.height())
            rx = max(0.0, min(1.0, rx))
            ry = max(0.0, min(1.0, ry))
            if rot:
                inv = (-rot) % 360
                rx, ry = self._rotate_norm_point(rx, ry, inv)
            start = int(getattr(b, "number", 0) or 0)
            end = int(getattr(b, "range_end", start) or start)
            try:
                bf = str(getattr(b, "backfill_rgb", "") or "").strip().upper()
            except Exception:
                bf = ""
            specs.append((start, end, float(rx), float(ry), int(b.base_radius), bf))

        specs.sort(key=lambda t: (t[0], t[1], t[2], t[3]))
        old_specs = self.bubble_specs_by_page.get(self.current_page, [])
        self.bubble_specs_by_page[self.current_page] = specs
        if old_specs != specs:
            self._set_dirty(True)
        if recompute_next:
            self._recompute_next_bubble_number()

        # Notify listeners (e.g. Form 3 bubble coloring) when bubble layout changes.
        if old_specs != specs:
            try:
                self.bubbles_changed.emit(self.get_bubbled_numbers())
            except Exception:
                pass

    def auto_resolve_bubble_overlaps_current_page(self, *, max_iters: int = 25) -> None:
        """Nudge bubbles left/right so they don't overlap.

        This runs after placing/resizing/renumbering/moving bubbles. It only adjusts
        X positions (left/right) and clamps bubbles within the rendered page.
        """

        if self.pixmap_item is None:
            return
        page_rect = self.pixmap_item.boundingRect()
        if page_rect.width() <= 0 or page_rect.height() <= 0:
            return

        bubbles = [b for b in (getattr(self, "bubbles", []) or []) if isinstance(b, BubbleItem)]
        if len(bubbles) < 2:
            return

        try:
            margin = max(2.0, float(getattr(self, "bubble_line_width", 3.0) or 3.0))
        except Exception:
            margin = 3.0

        def _scene_rect(b: BubbleItem) -> QRectF:
            try:
                return b.mapRectToScene(b.boundingRect())
            except Exception:
                # Fallback: approximate around position.
                try:
                    hw, hh = b._half_sizes()
                except Exception:
                    hw, hh = (20.0, 20.0)
                p = b.pos()
                return QRectF(float(p.x() - hw), float(p.y() - hh), float(hw * 2.0), float(hh * 2.0))

        def _clamp_x(b: BubbleItem) -> None:
            try:
                r = _scene_rect(b)
                dx = 0.0
                if r.left() < page_rect.left():
                    dx += float(page_rect.left() - r.left())
                if r.right() > page_rect.right():
                    dx -= float(r.right() - page_rect.right())
                if abs(dx) > 1e-3:
                    b.setPos(QPointF(float(b.pos().x() + dx), float(b.pos().y())))
            except Exception:
                return

        # Iteratively resolve overlaps.
        for _ in range(max(1, int(max_iters))):
            moved_any = False
            rects = [(b, _scene_rect(b)) for b in bubbles]
            rects.sort(key=lambda t: float(t[1].left()))

            for i in range(len(rects) - 1):
                b1, r1 = rects[i]
                for j in range(i + 1, len(rects)):
                    b2, r2 = rects[j]
                    # If the next rect starts after r1 ends, no further overlaps for this i.
                    if float(r2.left()) > float(r1.right()) + float(margin):
                        break
                    if not r1.intersects(r2):
                        continue

                    # Compute overlap in X (we only separate horizontally).
                    overlap_x = float(min(r1.right(), r2.right()) - max(r1.left(), r2.left()))
                    overlap_y = float(min(r1.bottom(), r2.bottom()) - max(r1.top(), r2.top()))
                    if overlap_x <= 0.0 or overlap_y <= 0.0:
                        continue

                    c1 = float(r1.center().x())
                    c2 = float(r2.center().x())

                    # If bubbles are stacked vertically (centers aligned within threshold),
                    # assume they are meant to be a column and DO NOT push them apart horizontally.
                    # This prevents vertical lists of bubbles (common in drawings) from "exploding" sideways.
                    vertical_align_threshold = min(r1.width(), r2.width()) * 0.25
                    if abs(c1 - c2) < vertical_align_threshold:
                        continue

                    # Push apart by half the overlap plus margin.
                    shift = (overlap_x / 2.0) + float(margin)
                    # Cap per-step shift to avoid wild jumps.
                    shift = min(shift, 120.0)

                    if c1 <= c2:
                        try:
                            b1.setPos(QPointF(float(b1.pos().x() - shift), float(b1.pos().y())))
                            b2.setPos(QPointF(float(b2.pos().x() + shift), float(b2.pos().y())))
                        except Exception:
                            pass
                    else:
                        try:
                            b1.setPos(QPointF(float(b1.pos().x() + shift), float(b1.pos().y())))
                            b2.setPos(QPointF(float(b2.pos().x() - shift), float(b2.pos().y())))
                        except Exception:
                            pass

                    _clamp_x(b1)
                    _clamp_x(b2)

                    moved_any = True
                    # Update cached rects for subsequent comparisons.
                    r1 = _scene_rect(b1)
                    rects[i] = (b1, r1)
                    rects[j] = (b2, _scene_rect(b2))

            if not moved_any:
                break

    def get_bubbled_numbers(self) -> set[int]:
        """Return all bubble numbers across all pages, expanding ranges."""
        out: set[int] = set()
        try:
            specs_by_page = getattr(self, "bubble_specs_by_page", {}) or {}
        except Exception:
            specs_by_page = {}

        for specs in specs_by_page.values():
            for spec in (specs or []):
                try:
                    start, end, _x, _y, _r = spec[:5]
                except Exception:
                    continue
                try:
                    s = int(start)
                    e = int(end)
                except Exception:
                    continue
                if s <= 0:
                    continue
                if e < s:
                    e = s
                # Safety cap to avoid pathological ranges.
                if e - s > 9999:
                    e = s
                for n in range(s, e + 1):
                    out.add(int(n))

        return out

    def find_bubble_page_index(self, number: int) -> int | None:
        """Return the 0-based page index containing the bubble number (or None)."""
        try:
            n = int(number)
        except Exception:
            return None
        if n <= 0:
            return None

        try:
            specs_by_page = getattr(self, "bubble_specs_by_page", {}) or {}
        except Exception:
            specs_by_page = {}

        for page_idx, specs in (specs_by_page or {}).items():
            try:
                p = int(page_idx)
            except Exception:
                continue
            for spec in (specs or []):
                try:
                    start, end, _x, _y, _r = spec[:5]
                except Exception:
                    continue
                try:
                    s = int(start)
                    e = int(end)
                except Exception:
                    continue
                if e < s:
                    e = s
                if s <= n <= e:
                    return int(p)

        return None

    def select_bubble_number(self, number: int, *, center: bool = True) -> bool:
        """Go to the bubble's page and select/highlight the bubble (including ranges).

        Returns True if a matching bubble/range was found and selected.
        """
        try:
            n = int(number)
        except Exception:
            return False
        if n <= 0:
            return False

        page_idx = self.find_bubble_page_index(int(n))
        if page_idx is None:
            return False

        # Navigate to the correct page.
        try:
            if int(page_idx) != int(getattr(self, "current_page", 0) or 0):
                try:
                    if hasattr(self, "page_combo") and self.page_combo is not None:
                        self.page_combo.blockSignals(True)
                        self.page_combo.setCurrentIndex(int(page_idx))
                        self.page_combo.blockSignals(False)
                except Exception:
                    pass
                try:
                    self.go_to_page(int(page_idx))
                except Exception:
                    pass
        except Exception:
            pass

        chosen = None
        try:
            for b in list(getattr(self, "bubbles", []) or []):
                try:
                    s = int(getattr(b, "number", 0) or 0)
                    e = int(getattr(b, "range_end", s) or s)
                except Exception:
                    continue
                if e < s:
                    e = s
                if s <= n <= e:
                    chosen = b
                    break
        except Exception:
            chosen = None

        if chosen is None:
            return False

        try:
            if getattr(self, "scene", None) is not None:
                try:
                    self.scene.clearSelection()
                except Exception:
                    for it in list(self.scene.selectedItems() or []):
                        try:
                            it.setSelected(False)
                        except Exception:
                            pass
        except Exception:
            pass

        try:
            chosen.setSelected(True)
        except Exception:
            pass

        if center:
            try:
                if getattr(self, "view", None) is not None:
                    self.view.centerOn(chosen)
                    try:
                        self.view.ensureVisible(chosen)
                    except Exception:
                        pass
            except Exception:
                pass

        try:
            chosen.update()
        except Exception:
            pass

        return True

    def apply_bubble_number_mapping(self, mapping: dict[int, int]) -> None:
        """Apply a renumber mapping (old -> new) to all bubble specs and visible items."""
        if not isinstance(mapping, dict) or not mapping:
            return

        # Normalize mapping to ints.
        norm: dict[int, int] = {}
        for k, v in (mapping or {}).items():
            try:
                ok = int(k)
                nv = int(v)
            except Exception:
                continue
            if ok > 0 and nv > 0 and ok != nv:
                norm[ok] = nv
        if not norm:
            return

        changed = False

        # Update stored specs for all pages.
        try:
            for page_idx, specs in list((getattr(self, "bubble_specs_by_page", {}) or {}).items()):
                new_specs: list[tuple[int, int, float, float, int, str]] = []
                page_changed = False
                for spec in (specs or []):
                    try:
                        start, end, x, y, r = spec[:5]
                        bf = spec[5] if len(spec) > 5 else ""
                    except Exception:
                        continue
                    try:
                        start, end, x, y, r = spec[:5]
                        bf = spec[5] if len(spec) > 5 else ""
                    except Exception:
                        new_specs.append((start, end, x, y, r, str(bf or "")))
                    try:
                        s = int(start)
                        e = int(end)
                    except Exception:
                        new_specs.append((start, end, x, y, r, str(bf or "")))
                    new_specs.append((int(s), int(e), float(x), float(y), int(r), str(bf or "")))
                    s2 = int(norm.get(s, s))
                    e2 = int(norm.get(e, e))
                    if s2 != s or e2 != e:
                        page_changed = True
                    # Keep order for ranges.
                    if e2 < s2:
                        e2 = s2
                    new_specs.append((int(s2), int(e2), float(x), float(y), int(r), str(bf or "")))
                if page_changed:
                    changed = True
                    new_specs.sort(key=lambda t: (t[0], t[1], t[2], t[3]))
                    self.bubble_specs_by_page[int(page_idx)] = new_specs
        except Exception:
            pass

        # Update visible items on the current page.
        try:
            for b in list(getattr(self, "bubbles", []) or []):
                try:
                    s = int(getattr(b, "number", 0) or 0)
                    e = int(getattr(b, "range_end", s) or s)
                except Exception:
                    continue
                s2 = int(norm.get(s, s))
                e2 = int(norm.get(e, e))
                if e2 < s2:
                    e2 = s2
                if s2 != s:
                    try:
                        b.set_number(int(s2))
                    except Exception:
                        pass
                    changed = True
                if e2 != e:
                    try:
                        b.set_range_end(int(e2))
                    except Exception:
                        pass
                    changed = True
        except Exception:
            pass

        if changed:
            try:
                self._set_dirty(True)
            except Exception:
                pass
            try:
                self._recompute_next_bubble_number()
            except Exception:
                pass
            try:
                self.bubbles_changed.emit(self.get_bubbled_numbers())
            except Exception:
                pass

    def delete_bubbles_with_numbers(self, numbers: set[int] | list[int] | tuple[int, ...]) -> None:
        """Delete single-number bubbles whose bubble number is in `numbers`.

        This is used by Form 3 row deletion: if a row is removed, the corresponding
        bubble (if it exists as a single bubble) should be removed from the drawing.
        Range bubbles are left intact.
        """
        try:
            targets = {int(n) for n in (numbers or []) if int(n) > 0}
        except Exception:
            targets = set()
        if not targets:
            return

        changed = False

        # Remove from stored specs (all pages).
        try:
            for page_idx, specs in list((getattr(self, "bubble_specs_by_page", {}) or {}).items()):
                new_specs: list[tuple[int, int, float, float, int, str]] = []
                page_changed = False
                for spec in (specs or []):
                    try:
                        start, end, x, y, r = spec[:5]
                        bf = spec[5] if len(spec) > 5 else ""
                    except Exception:
                        continue
                    try:
                        s = int(start)
                        e = int(end)
                    except Exception:
                        new_specs.append((start, end, x, y, r, str(bf or "")))
                        continue
                    # Only delete single bubbles (start==end) that match.
                    if s == e and s in targets:
                        page_changed = True
                        continue
                    new_specs.append((int(s), int(e), float(x), float(y), int(r), str(bf or "")))
                if page_changed:
                    changed = True
                    self.bubble_specs_by_page[int(page_idx)] = new_specs
        except Exception:
            pass

        # Remove visible items on the current page.
        try:
            for b in list(getattr(self, "bubbles", []) or []):
                try:
                    s = int(getattr(b, "number", 0) or 0)
                    e = int(getattr(b, "range_end", s) or s)
                except Exception:
                    continue
                if s == e and s in targets:
                    try:
                        self.remove_bubble(b)
                    except Exception:
                        pass
                    changed = True
        except Exception:
            pass

        if changed:
            try:
                self._set_dirty(True)
            except Exception:
                pass
            try:
                self._recompute_next_bubble_number()
            except Exception:
                pass
            try:
                self.bubbles_changed.emit(self.get_bubbled_numbers())
            except Exception:
                pass

    def exclude_numbers_from_ranges(self, numbers: set[int] | list[int] | tuple[int, ...]) -> None:
        """Remove specific bubble numbers from any range bubbles.

        This is used when inserting new Form 3 rows: if the inserted bubble number
        falls inside an existing range label (e.g. 14-28), it should *not* be treated
        as bubbled. We split the range into segments excluding that number
        (e.g. 14-16 and 18-28) and re-layout the segments to avoid overlap.
        """

        # Normalize input numbers.
        try:
            holes = sorted({int(n) for n in (numbers or []) if int(n) > 0})
        except Exception:
            holes = []
        if not holes:
            return

        try:
            specs_by_page = getattr(self, "bubble_specs_by_page", {}) or {}
        except Exception:
            specs_by_page = {}

        changed = False

        def _width_mult_for_label(label: str) -> float:
            n = len(str(label or ""))
            if n <= 2:
                return 1.0
            if n == 3:
                return 1.20
            if n == 4:
                return 1.35
            if n == 5:
                return 1.55
            if n == 6:
                return 1.75
            if n == 7:
                return 1.95
            if n <= 9:
                return 2.20
            return 2.50

        for page_idx, specs in list(specs_by_page.items()):
            try:
                page_i = int(page_idx)
            except Exception:
                continue
            cur_specs = list(specs or [])
            if not cur_specs:
                continue

            # Page width is needed to re-center split segments.
            page_w = None
            try:
                if getattr(self, "doc", None) is not None:
                    page = self.doc.load_page(int(page_i))
                    page_w = float(getattr(page, "rect", None).width)
            except Exception:
                page_w = None
            if not page_w or page_w <= 0:
                page_w = 1.0

            new_specs: list[tuple[int, int, float, float, int, str]] = []
            page_changed = False

            for spec in cur_specs:
                try:
                    start, end, rx, ry, br = spec[:5]
                    bf = spec[5] if len(spec) > 5 else ""
                except Exception:
                    continue
                try:
                    s = int(start)
                    e = int(end)
                except Exception:
                    new_specs.append((start, end, rx, ry, br, str(bf or "")))
                    continue
                if s <= 0:
                    new_specs.append((start, end, rx, ry, br, str(bf or "")))
                    continue
                if e < s:
                    e = s

                # Holes that apply to this spec.
                inside = [n for n in holes if s <= int(n) <= e]
                if not inside:
                    new_specs.append((int(s), int(e), float(rx), float(ry), int(br), str(bf or "")))
                    continue

                page_changed = True
                changed = True

                # If it's a single bubble and it's excluded, remove it.
                if s == e:
                    continue

                # Split into kept segments excluding holes.
                segments: list[tuple[int, int]] = []
                cur = int(s)
                for h in inside:
                    h = int(h)
                    if h < cur:
                        continue
                    if h > e:
                        break
                    if h > cur:
                        segments.append((int(cur), int(h - 1)))
                    cur = int(h + 1)
                if cur <= e:
                    segments.append((int(cur), int(e)))

                # If a split produced no segments, drop it.
                if not segments:
                    continue

                # If only one segment remains, keep it in the same position.
                if len(segments) == 1:
                    a, b = segments[0]
                    new_specs.append((int(a), int(b), float(rx), float(ry), int(br), str(bf or "")))
                    continue

                # Re-layout multiple segments side-by-side around the original center.
                cx_pts = float(rx) * float(page_w)
                pad_scene = max(6.0, float(br) * 0.6)
                pad_pts = float(pad_scene) / float(getattr(self, "base_render_scale", 2.0) or 2.0)

                labels = [f"{a}-{b}" if b > a else str(a) for (a, b) in segments]
                half_ws_pts: list[float] = []
                for lab in labels:
                    half_w_scene = float(br) * float(_width_mult_for_label(lab))
                    half_ws_pts.append(float(half_w_scene) / float(getattr(self, "base_render_scale", 2.0) or 2.0))

                total_w = 0.0
                for i in range(len(half_ws_pts)):
                    total_w += 2.0 * float(half_ws_pts[i])
                    if i != len(half_ws_pts) - 1:
                        total_w += float(pad_pts)

                left = float(cx_pts) - total_w / 2.0
                x_cursor = float(left)
                centers_pts: list[float] = []
                for i in range(len(half_ws_pts)):
                    hw = float(half_ws_pts[i])
                    centers_pts.append(float(x_cursor) + hw)
                    x_cursor += 2.0 * hw
                    if i != len(half_ws_pts) - 1:
                        x_cursor += float(pad_pts)

                for i, (a, b) in enumerate(segments):
                    try:
                        cx_i = float(centers_pts[i])
                    except Exception:
                        cx_i = float(cx_pts)
                    rx_i = max(0.0, min(1.0, float(cx_i) / float(page_w)))
                    new_specs.append((int(a), int(b), float(rx_i), float(ry), int(br)))

            if page_changed:
                new_specs.sort(key=lambda t: (t[0], t[1], t[2], t[3]))
                specs_by_page[int(page_i)] = new_specs

        if not changed:
            return

        try:
            self.bubble_specs_by_page = specs_by_page
        except Exception:
            pass

        try:
            self._set_dirty(True)
        except Exception:
            pass
        try:
            self._recompute_next_bubble_number()
        except Exception:
            pass

        # Re-render current page so bubbles visually update.
        # IMPORTANT: preserve the user's current zoom + view location.
        try:
            target_scale = float(getattr(self, "current_render_scale", 0.0) or 0.0)
            if target_scale <= 0.0:
                try:
                    target_scale = float(getattr(self, "base_render_scale", 2.0) or 2.0) * float(getattr(self, "_zoom_factor", 1.0) or 1.0)
                except Exception:
                    target_scale = 2.0
            src_rot = getattr(self, "_last_render_rotation", None)
            self.render_current_page(target_scale=target_scale, preserve_center=True, source_rotation=src_rot)
        except Exception:
            pass

        try:
            self.bubbles_changed.emit(self.get_bubbled_numbers())
        except Exception:
            pass

    def set_pending_bubble_number_to_lowest_available(self) -> None:
        """Set the UI's next bubble number to the lowest missing bubble number."""
        try:
            fixed = int(self._lowest_available_number())
        except Exception:
            fixed = int(getattr(self, "next_bubble_number", 1) or 1)
        try:
            self._set_pending_bubble_number(int(fixed))
        except Exception:
            pass

    def set_grid_enabled(self, enabled: bool) -> None:
        try:
            self.grid_enabled = bool(enabled)
        except Exception:
            self.grid_enabled = False
        try:
            self._settings.setValue("pdf_viewer/grid_enabled", bool(self.grid_enabled))
        except Exception:
            pass
        self._rebuild_grid_overlay()

    def set_grid_bounds_pct(self, left_pct: float, top_pct: float, width_pct: float, height_pct: float) -> None:
        try:
            self.grid_left_pct = float(left_pct)
            self.grid_top_pct = float(top_pct)
            self.grid_width_pct = float(width_pct)
            self.grid_height_pct = float(height_pct)
        except Exception:
            return
        try:
            self._settings.setValue("pdf_viewer/grid_left_pct", float(self.grid_left_pct))
            self._settings.setValue("pdf_viewer/grid_top_pct", float(self.grid_top_pct))
            self._settings.setValue("pdf_viewer/grid_width_pct", float(self.grid_width_pct))
            self._settings.setValue("pdf_viewer/grid_height_pct", float(self.grid_height_pct))
        except Exception:
            pass
        self._rebuild_grid_overlay()

        try:
            self.grid_bounds_changed.emit(
                float(self.grid_left_pct),
                float(self.grid_top_pct),
                float(self.grid_width_pct),
                float(self.grid_height_pct),
            )
        except Exception:
            pass

    def _grid_page_rect(self) -> QRectF | None:
        try:
            if self.pixmap_item is None:
                return None
            r = self.pixmap_item.boundingRect()
            return r if r.width() > 0 and r.height() > 0 else None
        except Exception:
            return None

    def _grid_rect_from_pcts(self) -> QRectF | None:
        page_rect = self._grid_page_rect()
        if page_rect is None:
            return None
        try:
            x0 = max(0.0, min(1.0, float(self.grid_left_pct) / 100.0))
            y0 = max(0.0, min(1.0, float(self.grid_top_pct) / 100.0))
            w = max(0.0, min(1.0, float(self.grid_width_pct) / 100.0))
            h = max(0.0, min(1.0, float(self.grid_height_pct) / 100.0))
        except Exception:
            x0, y0, w, h = 0.0, 0.0, 1.0, 1.0
        x1 = min(1.0, max(0.0, x0 + w))
        y1 = min(1.0, max(0.0, y0 + h))
        if x1 <= x0:
            x0, x1 = 0.0, 1.0
        if y1 <= y0:
            y0, y1 = 0.0, 1.0

        left = page_rect.left() + x0 * page_rect.width()
        top = page_rect.top() + y0 * page_rect.height()
        ww = (x1 - x0) * page_rect.width()
        hh = (y1 - y0) * page_rect.height()
        return QRectF(left, top, ww, hh)

    def _grid_pcts_from_rect(self, rect: QRectF) -> tuple[float, float, float, float]:
        page_rect = self._grid_page_rect()
        if page_rect is None:
            return (0.0, 0.0, 100.0, 100.0)

        try:
            x0 = (float(rect.left()) - float(page_rect.left())) / max(1e-9, float(page_rect.width()))
            y0 = (float(rect.top()) - float(page_rect.top())) / max(1e-9, float(page_rect.height()))
            x1 = (float(rect.right()) - float(page_rect.left())) / max(1e-9, float(page_rect.width()))
            y1 = (float(rect.bottom()) - float(page_rect.top())) / max(1e-9, float(page_rect.height()))
        except Exception:
            return (0.0, 0.0, 100.0, 100.0)

        x0 = max(0.0, min(1.0, x0))
        y0 = max(0.0, min(1.0, y0))
        x1 = max(0.0, min(1.0, x1))
        y1 = max(0.0, min(1.0, y1))
        if x1 < x0:
            x0, x1 = x1, x0
        if y1 < y0:
            y0, y1 = y1, y0

        left_pct = 100.0 * x0
        top_pct = 100.0 * y0
        width_pct = 100.0 * (x1 - x0)
        height_pct = 100.0 * (y1 - y0)
        return (left_pct, top_pct, width_pct, height_pct)

    def _update_grid_overlay_from_rect(self, rect: QRectF) -> None:
        """Update existing grid line/label items to match rect (no rebuild)."""
        try:
            cols = 8
            rows = 4
            cell_w = float(rect.width()) / float(cols)
            cell_h = float(rect.height()) / float(rows)
        except Exception:
            return

        # Lines
        try:
            for i, line in enumerate(list(self._grid_v_lines or [])):
                x = float(rect.left()) + float(i) * cell_w
                try:
                    line.setLine(x, float(rect.top()), x, float(rect.top()) + float(rect.height()))
                except Exception:
                    pass
        except Exception:
            pass

        try:
            for j, line in enumerate(list(self._grid_h_lines or [])):
                y = float(rect.top()) + float(j) * cell_h
                try:
                    line.setLine(float(rect.left()), y, float(rect.left()) + float(rect.width()), y)
                except Exception:
                    pass
        except Exception:
            pass

        # Labels
        try:
            idx = 0
            for r in range(rows):
                for c in range(cols):
                    if idx >= len(self._grid_label_items):
                        break
                    tx = float(rect.left()) + (float(c) + 0.5) * cell_w
                    ty = float(rect.top()) + (float(r) + 0.5) * cell_h
                    it = self._grid_label_items[idx]
                    br = it.boundingRect()
                    it.setPos(tx - br.width() / 2.0, ty - br.height() / 2.0)
                    idx += 1
        except Exception:
            pass

    def _on_grid_bounds_rect_dragging(self, rect: QRectF) -> None:
        """Called by the interactive bounds item during drag."""
        try:
            self._update_grid_overlay_from_rect(rect)
        except Exception:
            pass

        try:
            l, t, w, h = self._grid_pcts_from_rect(rect)
            self.grid_left_pct = float(l)
            self.grid_top_pct = float(t)
            self.grid_width_pct = float(w)
            self.grid_height_pct = float(h)
        except Exception:
            return

        try:
            self.grid_bounds_changed.emit(
                float(self.grid_left_pct),
                float(self.grid_top_pct),
                float(self.grid_width_pct),
                float(self.grid_height_pct),
            )
        except Exception:
            pass

    def _on_grid_bounds_rect_committed(self, rect: QRectF) -> None:
        """Called by the interactive bounds item on mouse release."""
        try:
            l, t, w, h = self._grid_pcts_from_rect(rect)
            self.set_grid_bounds_pct(float(l), float(t), float(w), float(h))
        except Exception:
            pass

    def _clear_grid_overlay(self) -> None:
        try:
            for it in list(self.grid_items or []):
                try:
                    self.scene.removeItem(it)
                except Exception:
                    pass
        except Exception:
            pass
        self.grid_items = []
        self._grid_v_lines = []
        self._grid_h_lines = []
        self._grid_label_items = []
        self._grid_bounds_item = None

    def _rebuild_grid_overlay(self) -> None:
        self._clear_grid_overlay()
        try:
            if not bool(getattr(self, "grid_enabled", False)):
                return
        except Exception:
            return
        if self.pixmap_item is None:
            return

        try:
            page_rect = self.pixmap_item.boundingRect()
        except Exception:
            return
        if page_rect.width() <= 0 or page_rect.height() <= 0:
            return

        try:
            x0 = max(0.0, min(1.0, float(self.grid_left_pct) / 100.0))
            y0 = max(0.0, min(1.0, float(self.grid_top_pct) / 100.0))
            w = max(0.0, min(1.0, float(self.grid_width_pct) / 100.0))
            h = max(0.0, min(1.0, float(self.grid_height_pct) / 100.0))
        except Exception:
            x0, y0, w, h = 0.0, 0.0, 1.0, 1.0
        x1 = min(1.0, max(0.0, x0 + w))
        y1 = min(1.0, max(0.0, y0 + h))
        if x1 <= x0:
            x0, x1 = 0.0, 1.0
        if y1 <= y0:
            y0, y1 = 0.0, 1.0

        grid_left = page_rect.left() + x0 * page_rect.width()
        grid_top = page_rect.top() + y0 * page_rect.height()
        grid_w = (x1 - x0) * page_rect.width()
        grid_h = (y1 - y0) * page_rect.height()

        cols = 8
        rows = 4
        cell_w = grid_w / float(cols)
        cell_h = grid_h / float(rows)

        try:
            pen = QPen(QColor(0, 0, 0, 160))
            pen.setWidthF(1.0)
        except Exception:
            pen = QPen(QColor(0, 0, 0, 160))

        # Vertical lines
        for i in range(cols + 1):
            x = grid_left + float(i) * cell_w
            try:
                line = QGraphicsLineItem(x, grid_top, x, grid_top + grid_h)
                line.setPen(pen)
                line.setZValue(50)
                self.scene.addItem(line)
                self.grid_items.append(line)
                self._grid_v_lines.append(line)
            except Exception:
                pass

        # Horizontal lines
        for j in range(rows + 1):
            y = grid_top + float(j) * cell_h
            try:
                line = QGraphicsLineItem(grid_left, y, grid_left + grid_w, y)
                line.setPen(pen)
                line.setZValue(50)
                self.scene.addItem(line)
                self.grid_items.append(line)
                self._grid_h_lines.append(line)
            except Exception:
                pass

        # Labels at cell centers
        row_labels = ["D", "C", "B", "A"]
        page_count = 0
        try:
            if self.doc:
                page_count = len(self.doc)
        except Exception:
            pass
        use_sheet_prefix = (page_count > 1)
        try:
            page_idx = int(self.current_page)
        except Exception:
            page_idx = 0

        for r in range(rows):
            for c in range(cols):
                try:
                    zone_num = 8 - c
                    zone_letter = row_labels[r]
                    label = f"{zone_letter}{zone_num}"
                    if use_sheet_prefix:
                        label = f"SH{page_idx + 1} {label}"
                    tx = grid_left + (c + 0.5) * cell_w
                    ty = grid_top + (r + 0.5) * cell_h
                    text_item = QGraphicsTextItem(label)
                    text_item.setDefaultTextColor(QColor(0, 0, 0, 180))
                    try:
                        f = QFont()
                        f.setPointSize(10)
                        text_item.setFont(f)
                    except Exception:
                        pass
                    br = text_item.boundingRect()
                    text_item.setPos(tx - br.width() / 2.0, ty - br.height() / 2.0)
                    text_item.setZValue(51)
                    self.scene.addItem(text_item)
                    self.grid_items.append(text_item)
                    self._grid_label_items.append(text_item)
                except Exception:
                    pass

        # Interactive bounds rectangle
        try:
            bounds_rect = QRectF(float(grid_left), float(grid_top), float(grid_w), float(grid_h))
            b = _GridBoundsItem(bounds_rect, viewer=self)
            self.scene.addItem(b)
            self.grid_items.append(b)
            self._grid_bounds_item = b
        except Exception:
            pass

    def get_bubble_zones(self) -> dict[int, str]:
        """
        Calculate the zone (e.g. 'SH1 A1') for each bubble.
        Returns a dict mapping bubble_number -> zone_string.
        Assumes standard grid: 8 columns (8..1 L->R), 4 rows (D..A T->B).
        """
        zones = {}
        page_count = 0
        try:
            if self.doc:
                page_count = len(self.doc)
        except Exception:
            pass
            
        use_sheet_prefix = (page_count > 1)
        
        # Rows: 0->D, 1->C, 2->B, 3->A
        row_labels = ["D", "C", "B", "A"]
        
        try:
            specs_by_page = getattr(self, "bubble_specs_by_page", {}) or {}
        except Exception:
            specs_by_page = {}

        # Grid bounds in normalized coordinates (0..1)
        try:
            x0 = max(0.0, min(1.0, float(self.grid_left_pct) / 100.0))
            y0 = max(0.0, min(1.0, float(self.grid_top_pct) / 100.0))
            w = max(0.0, min(1.0, float(self.grid_width_pct) / 100.0))
            h = max(0.0, min(1.0, float(self.grid_height_pct) / 100.0))
        except Exception:
            x0, y0, w, h = 0.0, 0.0, 1.0, 1.0
        x1 = min(1.0, max(0.0, x0 + w))
        y1 = min(1.0, max(0.0, y0 + h))
        if x1 <= x0:
            x0, x1 = 0.0, 1.0
        if y1 <= y0:
            y0, y1 = 0.0, 1.0

        # Estimate bubble footprint in normalized coords so bubbles that sit on
        # grid lines contribute to *all* touched zones.
        page_rect = None
        try:
            page_rect = self.pixmap_item.boundingRect() if self.pixmap_item is not None else None
        except Exception:
            page_rect = None
        try:
            pw = float(page_rect.width()) if page_rect is not None else 0.0
            ph = float(page_rect.height()) if page_rect is not None else 0.0
        except Exception:
            pw = ph = 0.0
        if pw <= 0 or ph <= 0:
            pw, ph = 1000.0, 1000.0

        def _zone_str_for_touched(page_idx: int, touched: set[tuple[int, int]]) -> str:
            # touched: {(row_idx, col_idx)}
            by_col: dict[int, set[str]] = {}
            for (ri, ci) in touched:
                try:
                    col_num = 8 - int(ci)
                    letter = str(row_labels[int(ri)])
                except Exception:
                    continue
                by_col.setdefault(int(col_num), set()).add(letter)

            parts: list[str] = []
            for col_num in sorted(by_col.keys()):
                letters = sorted(by_col[col_num])
                if not letters:
                    continue
                if len(letters) == 1:
                    parts.append(f"{letters[0]}{col_num}")
                else:
                    parts.append(f"{letters[0]}{col_num}-{letters[-1]}{col_num}")

            if not parts:
                return ""
            z = " ".join(parts)
            if use_sheet_prefix:
                return f"SH{int(page_idx) + 1} {z}"
            return z

        for page_idx, specs in specs_by_page.items():
            for spec in specs:
                try:
                    start, end, rx, ry, br = spec[:5]
                except Exception:
                    continue

                # Footprint in normalized coords.
                try:
                    rpx = max(0.0, float(int(br)))
                except Exception:
                    rpx = 0.0
                rx = float(rx)
                ry = float(ry)
                rnx = float(rpx) / max(1.0, pw)
                rny = float(rpx) / max(1.0, ph)

                bx0 = rx - rnx
                bx1 = rx + rnx
                by0 = ry - rny
                by1 = ry + rny

                # Convert to grid-relative normalized 0..1
                try:
                    gx0 = (bx0 - x0) / max(1e-9, (x1 - x0))
                    gx1 = (bx1 - x0) / max(1e-9, (x1 - x0))
                    gy0 = (by0 - y0) / max(1e-9, (y1 - y0))
                    gy1 = (by1 - y0) / max(1e-9, (y1 - y0))
                except Exception:
                    continue

                gx0, gx1 = sorted((gx0, gx1))
                gy0, gy1 = sorted((gy0, gy1))

                gx0 = max(0.0, min(0.999999, gx0))
                gx1 = max(0.0, min(0.999999, gx1))
                gy0 = max(0.0, min(0.999999, gy0))
                gy1 = max(0.0, min(0.999999, gy1))

                try:
                    c0 = int(gx0 * 8.0)
                    c1 = int(gx1 * 8.0)
                    r0 = int(gy0 * 4.0)
                    r1 = int(gy1 * 4.0)
                except Exception:
                    continue

                c0 = max(0, min(7, int(c0)))
                c1 = max(0, min(7, int(c1)))
                r0 = max(0, min(3, int(r0)))
                r1 = max(0, min(3, int(r1)))

                touched: set[tuple[int, int]] = set()
                for ri in range(min(r0, r1), max(r0, r1) + 1):
                    for ci in range(min(c0, c1), max(c0, c1) + 1):
                        touched.add((int(ri), int(ci)))

                zone_str = _zone_str_for_touched(int(page_idx), touched)
                if not zone_str:
                    continue

                try:
                    s = int(start)
                    e = int(end)
                    if e < s:
                        e = s
                    if e - s > 9999:
                        e = s
                except Exception:
                    continue
                for n in range(int(s), int(e) + 1):
                    zones[int(n)] = zone_str
                    
        return zones

    def get_reference_locations(self, mode: str | None = None) -> dict[int, str]:
        """Return reference location strings for bubbles based on mode.

        Modes:
          - 'sheet_zone': use Sheet/Zone (e.g. 'SH1 A1')
          - 'page_label': use Page label (e.g. 'PAGE 2')
          - 'none': return empty mapping
        """
        try:
            mode = str(mode or getattr(self, "reference_location_mode", "sheet_zone") or "sheet_zone").strip().lower()
        except Exception:
            mode = "sheet_zone"

        if mode in ("none", "off", "disable", "disabled"):
            return {}

        if mode in ("page", "page_label", "page_number", "page number"):
            locations: dict[int, str] = {}
            try:
                specs_by_page = getattr(self, "bubble_specs_by_page", {}) or {}
            except Exception:
                specs_by_page = {}

            for page_idx, specs in specs_by_page.items():
                try:
                    label = f"PAGE {int(page_idx) + 1}"
                except Exception:
                    label = "PAGE"
                for spec in list(specs):
                    try:
                        start, end, _rx, _ry, _r = spec[:5]
                    except Exception:
                        continue
                    try:
                        s = int(start)
                        e = int(end)
                        if e < s:
                            e = s
                        if e - s > 9999:
                            e = s
                        for n in range(s, e + 1):
                            locations[n] = label
                    except Exception:
                        continue
            return locations

        # Default to sheet/zone
        return self.get_bubble_zones()

    def _recompute_next_bubble_number(self) -> None:
        max_num = 0
        for specs in self.bubble_specs_by_page.values():
            for spec in specs:
                try:
                    _start, end, _x, _y, _r = spec[:5]
                except Exception:
                    continue
                if int(end) > max_num:
                    max_num = int(end)
        self.next_bubble_number = max(1, max_num + 1)
        try:
            if not bool(getattr(self, "placing_mode", False)) and not bool(getattr(self, "range_mode", False)):
                self._set_pending_bubble_number(self._lowest_available_number())
        except Exception:
            pass

    def toggle_range_mode(self, enabled: bool) -> None:
        self.range_mode = bool(enabled)
        self._range_end_number = None

        if enabled:
            # Disable other modes
            self.set_note_region_mode(False)
            self.add_bubble_btn.setChecked(False)
            self.add_bubble_btn.setText("Add Bubble")

            self.placing_mode = True
            self.view.set_placing_mode(True)
        else:
            if self.range_mode:
                self.range_mode = False
            self.placing_mode = False
            self.view.set_placing_mode(False)

    def open_add_range_dialog(self) -> None:
        """Legacy/compat: enable range mode; popup appears on canvas click."""
        try:
            self.add_range_btn.setChecked(True)
        except Exception:
            pass

    def _prompt_add_range(self, default_start: int) -> tuple[int, int] | None:
        """Prompt user for a start/end range; start is editable and must not overlap existing numbers."""
        default_start = max(1, min(9999, int(default_start)))

        while True:
            dlg = QDialog(self)
            dlg.setWindowTitle("Add Range")
            form = QFormLayout(dlg)

            start_spin = QSpinBox(dlg)
            start_spin.setRange(1, 9999)
            start_spin.setValue(int(default_start))

            end_spin = QSpinBox(dlg)
            end_spin.setRange(start_spin.value(), 9999)
            end_spin.setValue(int(default_start))

            def _sync_end_min(v: int) -> None:
                try:
                    end_spin.setMinimum(int(v))
                    if end_spin.value() < int(v):
                        end_spin.setValue(int(v))
                except Exception:
                    pass

            start_spin.valueChanged.connect(_sync_end_min)

            form.addRow("Start Number:", start_spin)
            form.addRow("End Number:", end_spin)

            buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, dlg)
            form.addRow(buttons)
            buttons.accepted.connect(dlg.accept)
            buttons.rejected.connect(dlg.reject)

            # Focus the End number field by default so user can type the range limit immediately.
            QTimer.singleShot(0, lambda: (end_spin.setFocus(), end_spin.selectAll()))

            if dlg.exec() != QDialog.Accepted:
                return None

            start = int(start_spin.value())
            end = int(end_spin.value())
            if end < start:
                end = start

            overlap = self._range_overlap(start, end)
            if overlap:
                shown = ", ".join(str(n) for n in overlap[:12])
                more = "" if len(overlap) <= 12 else f" (+{len(overlap) - 12} more)"
                QMessageBox.warning(
                    self,
                    "Duplicate Bubble Number",
                    f"One or more bubble numbers in {start}-{end} already exist: {shown}{more}.\n\nPlease choose a different range.",
                )
                default_start = start
                continue

            return (int(start), int(end))

    def _safe_filename_component(self, s: str) -> str:
        s = (s or "").strip()
        if not s:
            return ""
        s = re.sub(r"[\\/:*?\"<>|]", "_", s)
        s = re.sub(r"\s+", " ", s).strip()
        return s

    def _default_save_as_path(self) -> str:
        base = (self.default_save_basename or "").strip()
        if base:
            base = self._safe_filename_component(base)
            if not base.lower().endswith(".pdf"):
                base += ".pdf"
        else:
            seed = self._opened_pdf_path or self.file_path
            stem = os.path.splitext(os.path.basename(seed))[0]
            # Avoid repeatedly appending the suffix when saving multiple times.
            if stem.lower().endswith("_annotated"):
                base = stem + ".pdf"
            else:
                base = stem + "_annotated.pdf"

        default_dir = os.path.dirname(self._opened_pdf_path or self.file_path)
        return os.path.join(default_dir, base)

    def _save_drawing_to_path(self, out_path: str) -> bool:
        if not self.doc or not self.file_path:
            return False

        self._persist_current_page_bubbles(rotation=self._last_render_rotation)
        out_path = str(out_path or "").strip()
        if not out_path:
            return False

        try:
            base_pdf_path = self.file_path
            try:
                if not base_pdf_path or not os.path.exists(base_pdf_path):
                    base_pdf_path = self._opened_pdf_path or base_pdf_path
            except Exception:
                pass

            # Prefer cloning from the currently loaded in-memory document so any
            # embedded-annotation deletions are preserved.
            doc = None
            if self.doc is not None:
                try:
                    b = getattr(self, "_pdf_doc_bytes_after_annots_mutation", None)
                    if b:
                        if self._debug_enabled:
                            print("DEBUG: Using cached mutated bytes for save.")
                    else:
                        if self._debug_enabled:
                            print("DEBUG: Snapshotting current doc for save.")
                        b = self._snapshot_pdf_bytes(self.doc)
                    doc = fitz.open(stream=b, filetype="pdf")
                except Exception as e:
                    if self._debug_enabled:
                        print(f"DEBUG: Failed to open from bytes: {e}")
                    doc = None
            if doc is None:
                if self._debug_enabled:
                    print(f"DEBUG: Falling back to base PDF path: {base_pdf_path}")
                doc = fitz.open(base_pdf_path)

            # Ensure embedded-annotation deletions are reflected in the output.
            if bool(getattr(self, "_pdf_annots_mutated", False)):
                if self._debug_enabled:
                    print("DEBUG: Applying deleted annotations to output doc.")
                try:
                    self._apply_deleted_pdf_annotations_to_doc(doc)
                except Exception:
                    pass

            # CRITICAL: Remove any existing internal annotations from the output doc
            # before writing the new set. This prevents duplicates and ensures
            # that the "App Bubbles" are always fresh and correctly tagged.
            # ALSO: Remove "External" annotations that overlap with the new bubbles
            # (Collision Detection). This "claims" them so they don't show up as duplicates.
            try:
                for page_index in range(len(doc)):
                    try:
                        page = doc.load_page(page_index)
                        
                        # Pre-calculate centers of new bubbles for collision detection
                        new_bubble_centers = []
                        try:
                            page_w = page.rect.width
                            page_h = page.rect.height
                            specs = self.bubble_specs_by_page.get(page_index, [])
                            for spec in specs:
                                try:
                                    start, end, rx, ry, br = spec[:5]
                                except Exception:
                                    continue
                                cx = page_w * float(rx)
                                cy = page_h * float(ry)
                                new_bubble_centers.append((cx, cy))
                        except Exception:
                            pass

                        to_delete = []
                        for ann in (page.annots() or []):
                            try:
                                info = getattr(ann, "info", {}) or {}
                                is_internal = (info.get("title") == "AS9102_FAI_BUBBLE" or info.get("subject") == "AS9102_FAI_BUBBLE")
                                
                                if is_internal:
                                    # Always delete existing internal (we will rewrite them)
                                    to_delete.append(ann)
                                else:
                                    # Check for collision with new bubbles
                                    # If it collides, we assume it's the "Ext" version of the bubble we are saving, so delete it.
                                    try:
                                        ann_rect = ann.rect
                                        ann_cx = (ann_rect.x0 + ann_rect.x1) / 2.0
                                        ann_cy = (ann_rect.y0 + ann_rect.y1) / 2.0
                                        
                                        # Threshold: 15 points (approx 5mm) seems reasonable for "same bubble"
                                        threshold = 15.0 
                                        for (bcx, bcy) in new_bubble_centers:
                                            dx = abs(ann_cx - bcx)
                                            dy = abs(ann_cy - bcy)
                                            if dx < threshold and dy < threshold:
                                                to_delete.append(ann)
                                                break
                                    except Exception:
                                        pass

                            except Exception:
                                pass
                        for ann in to_delete:
                            try:
                                page.delete_annot(ann)
                            except Exception:
                                pass
                    except Exception:
                        pass
            except Exception:
                pass

            def _width_mult_for_label(label: str) -> float:
                n = len(str(label or ""))
                if n <= 2:
                    return 1.0
                if n == 3:
                    return 1.20
                if n == 4:
                    return 1.35
                if n == 5:
                    return 1.55
                if n == 6:
                    return 1.75
                if n == 7:
                    return 1.95
                if n <= 9:
                    return 2.20
                return 2.50

            # Export as *real* PDF annotations so external editors can select/copy/paste.
            bubble_color = (0.86, 0.16, 0.16)
            try:
                qc = getattr(self, "bubble_color", None)
                if qc is not None and qc.isValid():
                    bubble_color = (qc.redF(), qc.greenF(), qc.blueF())
            except Exception:
                pass
            try:
                center_align = int(getattr(fitz, "TEXT_ALIGN_CENTER", 1))
            except Exception:
                center_align = 1

            def _freetext_rect_and_font(cx: float, cy: float, radius_pts: float, lw_pts: float, label: str) -> tuple[fitz.Rect, float]:
                """Return (rect, fontsize) tuned so the label fits."""
                label = str(label or "")
                n = max(1, len(label))

                half_h = float(radius_pts)
                # Start from the existing width multiplier logic.
                half_w = float(radius_pts) * float(_width_mult_for_label(label))

                # Padding inside the border.
                pad = max(2.0, float(lw_pts) * 1.5)

                # Use a simple width model for Helvetica-ish fonts:
                # average character width ~= 0.60 * fontsize.
                avg_char_w = 0.60
                min_font = 6.0

                # Prefer a size that matches the on-screen look (mainly driven by height),
                # and *widen the box* for long labels instead of shrinking font size.
                preferred = max(min_font, half_h * 1.05)

                inner_h = max(1.0, 2.0 * half_h - 2.0 * pad)
                # Height model: text height ~= 1.2 * fontsize.
                max_font_by_h = inner_h / 1.2
                fontsize = max(min_font, min(preferred, max_font_by_h))

                # Ensure width can fit this fontsize without wrapping in Kofax.
                # Add a small safety margin to avoid viewer-side metric differences clipping text.
                need_half_w = (avg_char_w * float(n) * float(fontsize) + 2.0 * pad) / 2.0
                need_half_w *= 1.08
                if half_w < need_half_w:
                    half_w = need_half_w

                rect = fitz.Rect(cx - half_w, cy - half_h, cx + half_w, cy + half_h)
                return rect, float(fontsize)


            def _add_outline_annot(page: fitz.Page, rect: fitz.Rect, shape: str, lw_pts: float, fill_rgb: str | None = None) -> None:
                """Add a visible outline annotation that Kofax reliably renders."""
                shape = str(shape or "Circle")
                try:
                    if shape.lower().startswith("rect"):
                        outline = page.add_rect_annot(rect)
                    else:
                        outline = page.add_circle_annot(rect)
                except Exception:
                    return
                try:
                    fill = None
                    rgb = str(fill_rgb or "").strip()
                    if not rgb:
                        if bool(getattr(self, "bubble_backfill_white", False)):
                            try:
                                rgb = getattr(self, "bubble_backfill_rgb", "FFFFFF")
                            except Exception:
                                rgb = "FFFFFF"
                    rgb = re.sub(r"[^0-9a-fA-F]", "", str(rgb))
                    if len(rgb) == 6:
                        try:
                            r = int(rgb[0:2], 16) / 255.0
                            g = int(rgb[2:4], 16) / 255.0
                            b = int(rgb[4:6], 16) / 255.0
                            fill = (r, g, b)
                        except Exception:
                            fill = (1.0, 1.0, 1.0)
                    outline.set_colors(stroke=bubble_color, fill=fill)
                except Exception:
                    pass
                try:
                    # Kofax can render very thin borders poorly; enforce a slightly thicker minimum.
                    outline.set_border(width=max(1.0, float(lw_pts)))
                except Exception:
                    pass
                try:
                    # Tag as internal annotation
                    outline.set_info({"title": "AS9102_FAI_BUBBLE", "subject": "AS9102_FAI_BUBBLE"})
                except Exception:
                    pass
                try:
                    outline.update()
                except Exception:
                    pass

            for page_index in range(len(doc)):
                page = doc.load_page(page_index)
                page_rect = page.rect
                specs = self.bubble_specs_by_page.get(page_index, [])
                if not specs:
                    continue

                for spec in specs:
                    try:
                        start, end, rx, ry, br = spec[:5]
                        bf = spec[5] if len(spec) > 5 else ""
                    except Exception:
                        continue
                    cx = page_rect.width * float(rx)
                    cy = page_rect.height * float(ry)
                    radius_pts = float(br) / float(self.base_render_scale)
                    lw_pts = float(self.bubble_line_width) / float(self.base_render_scale)

                    label = f"{int(start)}-{int(end)}" if int(end) > int(start) else str(int(start))
                    rect, font_size = _freetext_rect_and_font(cx, cy, radius_pts, lw_pts, label)

                    # Kofax may ignore custom FreeText /AP streams. Use a real shape annotation
                    # for the bubble outline so the circle/rectangle is always visible.
                    _add_outline_annot(page, rect, str(getattr(self, "bubble_shape", "Circle")), float(lw_pts), str(bf or ""))

                    # Use FreeText annotation (matches the example "Formal Document" bubbles).
                    # Many PDF tools store FreeText 'content' with a trailing CR.
                    # Matching this tends to improve baseline/centering consistency across viewers.
                    ann = page.add_freetext_annot(
                        rect,
                        # Avoid CR padding tricks; keep single-line content so viewers don't
                        # apply inconsistent multi-line vertical alignment.
                        str(label),
                        fontsize=float(font_size),
                        fontname=None,
                        text_color=bubble_color,
                        fill_color=None,
                        align=center_align,
                        opacity=1,
                        rotate=0,
                        richtext=False,
                    )

                    # Kofax often renders FreeText content top-aligned when regenerating the
                    # appearance. Setting /RD (inner padding) helps force the text line to be
                    # centered within the annotation rectangle across viewers.
                    try:
                        rect_h = float(rect.height)
                        rect_w = float(rect.width)
                        # One line text height approximation.
                        line_h = 1.20 * float(font_size)
                        pad_v = max(0.0, (rect_h - line_h) / 2.0)
                        # Keep horizontal padding small; large /RD left-right padding can clip
                        # long labels in some viewers (Kofax in particular).
                        pad_h = max(1.0, min(3.0, 0.25 * float(font_size)))
                        doc.xref_set_key(
                            ann.xref,
                            "RD",
                            f"[{pad_h:.3f} {pad_v:.3f} {pad_h:.3f} {pad_v:.3f}]",
                        )
                    except Exception:
                        pass
                    try:
                        # Avoid a rectangular FreeText border; outline comes from the shape annot.
                        ann.set_border(width=0)
                    except Exception:
                        pass
                    try:
                        ann.set_info({"title": "AS9102_FAI_BUBBLE", "subject": "AS9102_FAI_BUBBLE"})
                    except Exception:
                        pass
                    try:
                        ann.update()
                    except Exception:
                        pass

            # Full rewrite so deleted embedded annotations cannot reappear.
            # garbage=0 is safest for metadata persistence and avoiding xref errors.
            doc.save(out_path, garbage=0, deflate=True)
            doc.close()

            # Save editable state alongside the output PDF so reopening stays editable.
            try:
                self._save_edit_state_sidecar(out_path)
            except Exception:
                pass

            self._set_dirty(False)
            self._last_saved_pdf_path = str(out_path)

            # Treat the saved PDF as the current opened drawing path for subsequent operations.
            self._opened_pdf_path = str(out_path)
            self._sidecar_context_pdf_path = str(out_path)

            # Traditional Save: subsequent saves go to the same target without prompting.
            self._save_target_pdf_path = str(out_path)

            try:
                self.drawing_saved.emit(str(out_path))
            except Exception:
                pass
            
            # Disable auto-import after a successful save to prevent re-importing
            # external bubbles that might have been baked in or are no longer needed.
            try:
                self.auto_import_annots = False
                # If we have a button for this in the UI (via DrawingViewerWindow), update it.
                # Since we don't have a direct reference to the button here, we rely on
                # the button checking the attribute or a signal if we had one.
                # However, DrawingViewerWindow connects the button to setattr.
                # We can try to emit a signal or just let the UI update on next refresh if bound.
                # For now, just setting the flag is the logic change requested.
            except Exception:
                pass
                
            return True
        except Exception as e:
            QMessageBox.critical(self, "Save Failed", f"Failed to save annotated drawing:\n{e}")
            return False

    def save_drawing_as(self) -> bool:
        """Save As... always prompts for a file path."""
        if not self.doc or not self.file_path:
            return False

        default_path = self._default_save_as_path()
        out_path, _ = QFileDialog.getSaveFileName(self, "Save Drawing As", default_path, "PDF Files (*.pdf)")
        if not out_path:
            return False
        return self._save_drawing_to_path(out_path)

    def save_drawing(self) -> bool:
        """Save using the current save target; if none, behaves like Save As."""
        target = str(getattr(self, "_save_target_pdf_path", "") or "").strip()
        if not target:
            return self.save_drawing_as()
        return self._save_drawing_to_path(target)

    def add_bubble(self, number, x, y):
        """Add a bubble."""
        self._push_undo_state()
        self._add_bubble_internal(number, x, y, self.bubble_base_radius)
        
        if number >= self.next_bubble_number:
            self.next_bubble_number = number + 1
        
        self.bubble_added.emit(number)

        self._set_dirty(True)

    def add_range_bubble(self, start: int, end: int, x: float, y: float) -> None:
        self._push_undo_state()
        start = int(start)
        end = int(end)
        if end < start:
            end = start
        label = f"{start}-{end}" if end > start else str(start)
        bubble = BubbleItem(
            start,
            x,
            y,
            base_radius=self.bubble_base_radius,
            parent_viewer=self,
            range_end=end,
            display_text=label,
            backfill_rgb="FFFFFF",
        )
        # The placement click is forwarded to the scene; suppress undo snapshot for that immediate press.
        bubble._suppress_next_press_undo = True
        self.scene.addItem(bubble)
        self.bubbles.append(bubble)

        # Advance next bubble number past the range.
        self.next_bubble_number = max(self.next_bubble_number, end + 1)
        try:
            self.auto_resolve_bubble_overlaps_current_page()
        except Exception:
            pass
        self._persist_current_page_bubbles()

        self._set_dirty(True)

    def _add_bubble_internal(self, number, x, y, radius):
        bubble = BubbleItem(number, x, y, base_radius=radius, parent_viewer=self, backfill_rgb="FFFFFF")
        # The placement click is forwarded to the scene; suppress undo snapshot for that immediate press.
        bubble._suppress_next_press_undo = True
        self.scene.addItem(bubble)
        self.bubbles.append(bubble)
        try:
            self.auto_resolve_bubble_overlaps_current_page()
        except Exception:
            pass
        self._persist_current_page_bubbles()

    def add_bubble_from_drop(self, text, x, y):
        self.add_bubble(self.next_bubble_number, x, y)

    def remove_bubble(self, bubble):
        if bubble in self.bubbles:
            self._push_undo_state()
            self.bubbles.remove(bubble)
            self.scene.removeItem(bubble)
            try:
                self.auto_resolve_bubble_overlaps_current_page()
            except Exception:
                pass
            self._persist_current_page_bubbles()
            self.bubble_removed.emit(bubble.number)

            self._set_dirty(True)

    def renumber_bubble(self, bubble):
        start_current = int(getattr(bubble, "number", 1) or 1)
        end_current = int(getattr(bubble, "range_end", start_current) or start_current)
        is_range = end_current > start_current

        if is_range:
            dlg = QDialog(self)
            dlg.setWindowTitle("Renumber Range")
            form = QFormLayout(dlg)

            start_spin = QSpinBox(dlg)
            start_spin.setRange(1, 9999)
            start_spin.setValue(start_current)

            end_spin = QSpinBox(dlg)
            end_spin.setRange(start_spin.value(), 9999)
            end_spin.setValue(max(end_current, start_spin.value()))

            def _sync_end_min(v: int) -> None:
                try:
                    end_spin.setMinimum(int(v))
                    if end_spin.value() < int(v):
                        end_spin.setValue(int(v))
                except Exception:
                    pass

            start_spin.valueChanged.connect(_sync_end_min)

            form.addRow("Start Number:", start_spin)
            form.addRow("End Number:", end_spin)

            buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, dlg)
            buttons.accepted.connect(dlg.accept)
            buttons.rejected.connect(dlg.reject)
            form.addRow(buttons)

            if dlg.exec() != QDialog.Accepted:
                return

            new_start = int(start_spin.value())
            new_end = int(end_spin.value())
            if new_end < new_start:
                new_end = new_start

            self._push_undo_state()
            bubble.set_number(new_start)
            bubble.set_range_end(new_end)
            try:
                self.auto_resolve_bubble_overlaps_current_page()
            except Exception:
                pass
            self._persist_current_page_bubbles()
            self._set_dirty(True)
            return

        number, ok = QInputDialog.getInt(self, "Renumber", "New Number:", start_current, 1, 9999)
        if ok:
            self._push_undo_state()
            bubble.set_number(number)
            try:
                self.auto_resolve_bubble_overlaps_current_page()
            except Exception:
                pass
            self._persist_current_page_bubbles()
            if number >= self.next_bubble_number:
                self.next_bubble_number = number + 1
            self._set_dirty(True)

    def resize_bubble(self, bubble):
        current_scale = self._size_slider_from_radius(getattr(bubble, "base_radius", self.bubble_base_radius))
        scale, ok = QInputDialog.getInt(self, "Resize", "Size (1-10):", 
                                        current_scale, 1, 10)
        if ok:
            self._push_undo_state()
            new_radius = self._radius_from_size_slider(scale)
            bubble.set_base_radius(new_radius)
            self.bubble_base_radius = new_radius
            self.size_slider.setValue(scale)
            try:
                self.auto_resolve_bubble_overlaps_current_page()
            except Exception:
                pass
            self._persist_current_page_bubbles()
            self._set_dirty(True)

    def clear_bubbles(self):
        # Clear bubbles on the CURRENT page only.
        try:
            existing = list(getattr(self, "bubbles", []) or [])
        except Exception:
            existing = []

        if existing:
            try:
                mb = QMessageBox(self)
                mb.setIcon(QMessageBox.Warning)
                mb.setWindowTitle("Clear Bubbles")
                mb.setText("Clear all bubbles on the current page?\n\nThis cannot be undone.")
                mb.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
                mb.setDefaultButton(QMessageBox.No)
                if mb.exec() != QMessageBox.Yes:
                    return
            except Exception:
                pass

        self._push_undo_state()
        for bubble in self.bubbles[:]:
            self.scene.removeItem(bubble)
        self.bubbles = []
        self.bubble_specs_by_page[self.current_page] = []
        self._recompute_next_bubble_number()

        self._set_dirty(True)
        
        if self.placing_mode:
            self.add_bubble_btn.setText(f"Click to place #{self.next_bubble_number}")

    def handle_file_drop(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        if files and files[0].lower().endswith('.pdf'):
            self.load_pdf(files[0])
            event.accept()

    def showEvent(self, event):
        super().showEvent(event)
        if not self._did_initial_fit:
            QTimer.singleShot(100, self.fit_to_view)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.on_escape()
        else:
            super().keyPressEvent(event)

