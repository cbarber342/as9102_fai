from __future__ import annotations

import os

from PySide6.QtCore import Qt, Signal, QSettings, QTimer
from PySide6.QtGui import QPalette, QColor
from PySide6.QtWidgets import (
    QDockWidget,
    QFormLayout,
    QHBoxLayout,
    QCheckBox,
    QButtonGroup,
    QColorDialog,
    QGroupBox,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QDoubleSpinBox,
    QSizePolicy,
    QToolBar,
    QTabWidget,
    QVBoxLayout,
    QWidget,
    QComboBox,
)

from as9102_fai.gui.pdf_viewer import PdfViewer


class _ColorSwatchCheckBox(QCheckBox):
    """A checkbox rendered as a color swatch that supports right-click color picking."""

    colorChanged = Signal()

    def __init__(self, label: str = "", rgb: str | None = None, parent: QWidget | None = None):
        super().__init__(label, parent)
        self._swatch_rgb: str | None = None
        self.setFixedSize(22, 22)
        self.set_swatch_rgb(rgb)

    def swatch_rgb(self) -> str | None:
        return self._swatch_rgb

    def set_swatch_rgb(self, rgb: str | None) -> None:
        s = str(rgb or "").strip().lstrip("#")
        self._swatch_rgb = s.upper() if s else None
        bg = f"#{self._swatch_rgb}" if self._swatch_rgb else "palette(base)"
        self.setStyleSheet(
            "QCheckBox::indicator { width: 18px; height: 18px; }"
            f"QCheckBox::indicator:unchecked {{ background: {bg}; border: 1px solid #666; }}"
            f"QCheckBox::indicator:checked {{ background: {bg}; border: 2px solid #000; }}"
        )

    def mousePressEvent(self, event):
        try:
            if event.button() == Qt.MouseButton.RightButton:
                menu = QMenu(self)
                act_pick = menu.addAction("Choose color…")
                act_none = menu.addAction("No color")
                chosen = menu.exec(event.globalPosition().toPoint() if hasattr(event, "globalPosition") else event.globalPos())

                if chosen == act_none:
                    self.set_swatch_rgb(None)
                    try:
                        self.colorChanged.emit()
                    except Exception:
                        pass
                elif chosen == act_pick:
                    cur = QColor("#" + self._swatch_rgb) if self._swatch_rgb else QColor()
                    color = QColorDialog.getColor(cur, self, "Select Color")
                    if color.isValid():
                        self.set_swatch_rgb(color.name()[1:])
                        try:
                            self.colorChanged.emit()
                        except Exception:
                            pass
                event.accept()
                return
        except Exception:
            pass
        return super().mousePressEvent(event)


class DrawingViewerWindow(QMainWindow):
    bubbleScrollSelected = Signal(int, int)  # (start, end)

    def __init__(
        self,
        *,
        pdf_path: str = "",
        default_save_basename: str = "",
        pop_out_callback=None,
        dock_back_callback=None,
    ):
        super().__init__()
        self.current_pdf_path = str(pdf_path)
        self._pop_out_callback = pop_out_callback
        self._dock_back_callback = dock_back_callback
        self._settings = QSettings()
        self._pop_btn = None
        self._dock_btn = None
        self._pdf_viewer = PdfViewer(default_save_basename=default_save_basename, embed_controls=False)

        # Defer loading if path is empty/invalid (allows embedding in main tabs).
        try:
            if pdf_path and os.path.exists(str(pdf_path)):
                self._pdf_viewer.load_pdf(pdf_path)
        except Exception:
            pass

        try:
            self._pdf_viewer.drawing_saved.connect(self._on_drawing_saved)
        except Exception:
            pass

        self.setWindowTitle("Drawing Viewer")
        self.setCentralWidget(self._pdf_viewer)

        self._build_toolbar()
        self._build_docks()
        self._build_status_bar()

    def _apply_theme_aware_combobox_style(self, combo: QComboBox) -> None:
        """Make combo popup readable (avoid transparent popup in some styles)."""
        if combo is None:
            return

        # Use palette roles so this auto-updates on theme switches.
        try:
            combo.setStyleSheet(
                "QComboBox { background-color: palette(base); color: palette(text); }"
                "QComboBox QAbstractItemView {"
                "  background-color: palette(base);"
                "  color: palette(text);"
                "  selection-background-color: palette(highlight);"
                "  selection-color: palette(highlighted-text);"
                "}"
            )
        except Exception:
            pass

        # Some styles need the popup view to be non-transparent.
        try:
            v = combo.view()
            if v is not None:
                v.setAutoFillBackground(True)
        except Exception:
            pass

    def load_pdf(self, pdf_path: str) -> None:
        p = str(pdf_path or "").strip()
        if not p:
            return
        try:
            if not os.path.exists(p):
                return
        except Exception:
            return

        try:
            self.current_pdf_path = p
        except Exception:
            pass

        try:
            if getattr(self._pdf_viewer, "_debug_enabled", False):
                print(f"[AS9102_DEBUG_PDF] DrawingViewerWindow.load_pdf: {p}", flush=True)
        except Exception:
            pass

        try:
            self._pdf_viewer.load_pdf(p)
        except Exception:
            return

        try:
            self.setWindowTitle(f"Drawing Viewer - {p}")
        except Exception:
            pass

    def _on_drawing_saved(self, out_path: str) -> None:
        self.current_pdf_path = str(out_path or "")
        try:
            name = self.current_pdf_path
            if name:
                self.setWindowTitle(f"Drawing Viewer - {name}")
        except Exception:
            pass

    def closeEvent(self, event):
        # If this viewer is currently undocked (Dock Back visible), closing the
        # window should re-dock it instead of destroying the widget.
        try:
            if self._dock_back_callback is not None and self._dock_btn is not None and self._dock_btn.isVisible():
                self._dock_back_callback()
                event.ignore()
                return
        except Exception:
            pass

        try:
            if self._pdf_viewer is not None and not self._pdf_viewer.can_close():
                event.ignore()
                return
        except Exception:
            pass

        super().closeEvent(event)

    def _build_toolbar(self) -> None:
        tb = QToolBar("Drawing")
        tb.setMovable(True)
        tb.setFloatable(False)
        self.addToolBar(Qt.TopToolBarArea, tb)

        v = self._pdf_viewer
        tb.addWidget(v.add_bubble_btn)
        try:
            tb.addWidget(v.bubble_number_spin)
        except Exception:
            pass
        tb.addWidget(v.add_range_btn)
        tb.addSeparator()
        tb.addWidget(v.save_drawing_btn)
        tb.addWidget(v.save_drawing_as_btn)

        tb.addSeparator()

        # Bubble backfill swatches (apply to selected bubbles)
        try:
            tb.addWidget(QLabel("Backfill:"))
            bf1 = str(self._settings.value("pdf_viewer/bubble_backfill_swatch1_rgb", "FFFFFF", type=str) or "").strip()
            bf2 = str(self._settings.value("pdf_viewer/bubble_backfill_swatch2_rgb", "FFFFFF", type=str) or "").strip()
            bf3 = str(self._settings.value("pdf_viewer/bubble_backfill_swatch3_rgb", "FFFFFF", type=str) or "").strip()

            bf_btn1 = _ColorSwatchCheckBox("", bf1 or None, tb)
            bf_btn2 = _ColorSwatchCheckBox("", bf2 or None, tb)
            bf_btn3 = _ColorSwatchCheckBox("", bf3 or None, tb)

            bf_btn1.setToolTip("Backfill color 1 (right-click to change)")
            bf_btn2.setToolTip("Backfill color 2 (right-click to change)")
            bf_btn3.setToolTip("Backfill color 3 (default white; right-click to change)")

            bf_group = QButtonGroup(tb)
            bf_group.setExclusive(True)
            bf_group.addButton(bf_btn1, 1)
            bf_group.addButton(bf_btn2, 2)
            bf_group.addButton(bf_btn3, 3)

            def _clear_bf_selection() -> None:
                try:
                    bf_group.setExclusive(False)
                    bf_btn1.setChecked(False)
                    bf_btn2.setChecked(False)
                    bf_btn3.setChecked(False)
                finally:
                    bf_group.setExclusive(True)

            def _persist_bf_swatches() -> None:
                try:
                    self._settings.setValue("pdf_viewer/bubble_backfill_swatch1_rgb", bf_btn1.swatch_rgb() or "")
                    self._settings.setValue("pdf_viewer/bubble_backfill_swatch2_rgb", bf_btn2.swatch_rgb() or "")
                    self._settings.setValue("pdf_viewer/bubble_backfill_swatch3_rgb", bf_btn3.swatch_rgb() or "")
                except Exception:
                    pass

            tb.addWidget(bf_btn1)
            tb.addWidget(bf_btn2)
            tb.addWidget(bf_btn3)

            def _apply_bf_from_swatch() -> None:
                checked_id = -1
                try:
                    checked_id = int(bf_group.checkedId())
                except Exception:
                    checked_id = -1

                if checked_id == 1:
                    rgb = bf_btn1.swatch_rgb() or ""
                elif checked_id == 2:
                    rgb = bf_btn2.swatch_rgb() or ""
                elif checked_id == 3:
                    rgb = bf_btn3.swatch_rgb() or ""
                else:
                    return

                applied = False
                try:
                    if hasattr(self._pdf_viewer, "apply_backfill_to_selected_bubbles"):
                        applied = bool(self._pdf_viewer.apply_backfill_to_selected_bubbles(rgb))
                except Exception:
                    applied = False

                if not applied:
                    try:
                        QMessageBox.information(self, "Bubble Backfill", "Select one or more bubbles first.")
                    except Exception:
                        pass

                try:
                    QTimer.singleShot(0, _clear_bf_selection)
                except Exception:
                    pass

            bf_group.buttonClicked.connect(lambda _b=None: _apply_bf_from_swatch())
            bf_btn1.colorChanged.connect(_persist_bf_swatches)
            bf_btn2.colorChanged.connect(_persist_bf_swatches)
            bf_btn3.colorChanged.connect(_persist_bf_swatches)
        except Exception:
            pass

        try:
            tb.addWidget(QLabel("Find bubble:"))
            self._find_bubble_edit = QLineEdit()
            self._find_bubble_edit.setPlaceholderText("#")
            self._find_bubble_edit.setMaximumWidth(70)
            tb.addWidget(self._find_bubble_edit)

            self._find_bubble_btn = QPushButton("Go")
            self._find_bubble_btn.setMaximumWidth(48)
            tb.addWidget(self._find_bubble_btn)

            def _do_find() -> None:
                try:
                    n = int(str(self._find_bubble_edit.text() or "").strip())
                except Exception:
                    n = 0
                if n <= 0:
                    QMessageBox.information(self, "Find Bubble", "Enter a valid bubble number.")
                    return
                try:
                    if hasattr(self._pdf_viewer, "select_bubble_number"):
                        ok = bool(self._pdf_viewer.select_bubble_number(int(n), center=True))
                    else:
                        ok = False
                except Exception:
                    ok = False
                if not ok:
                    QMessageBox.information(self, "Find Bubble", f"Bubble {n} was not found on the drawing.")


                # Keep the Bubble scroller in sync with Find Bubble.
                try:
                    setter = getattr(self, "_bubble_scroller_set_to_number", None)
                    if callable(setter):
                        setter(int(n))
                except Exception:
                    pass
            self._find_bubble_btn.clicked.connect(lambda _checked=False: _do_find())
            self._find_bubble_edit.returnPressed.connect(_do_find)
        except Exception:
            self._find_bubble_edit = None
            self._find_bubble_btn = None

        # Bubble scroller: step through existing bubble *items* (including ranges).
        try:
            tb.addSeparator()
            tb.addWidget(QLabel("Bubble:"))

            self._bubble_entries: list[tuple[int, int, str]] = []  # (start, end, label)
            self._bubble_entry_index: int = -1

            self._bubble_scroll_edit = QLineEdit()
            self._bubble_scroll_edit.setReadOnly(True)
            self._bubble_scroll_edit.setMaximumWidth(80)
            self._bubble_scroll_edit.setPlaceholderText("-")
            tb.addWidget(self._bubble_scroll_edit)

            self._bubble_scroll_up = QPushButton("▲")
            self._bubble_scroll_up.setMaximumWidth(28)
            self._bubble_scroll_down = QPushButton("▼")
            self._bubble_scroll_down.setMaximumWidth(28)
            tb.addWidget(self._bubble_scroll_up)
            tb.addWidget(self._bubble_scroll_down)

            def _refresh_bubble_entries() -> None:
                entries: list[tuple[int, int, str]] = []
                try:
                    specs_by_page = getattr(self._pdf_viewer, "bubble_specs_by_page", {}) or {}
                except Exception:
                    specs_by_page = {}

                for _page_idx, specs in (specs_by_page or {}).items():
                    for spec in (specs or []):
                        # bubble_specs may include additional fields (color, rect, etc.).
                        try:
                            start = spec[0]
                            end = spec[1]
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
                        label = f"{s}-{e}" if e > s else str(s)
                        entries.append((int(s), int(e), str(label)))

                # Deduplicate by (start,end)
                uniq: dict[tuple[int, int], tuple[int, int, str]] = {}
                for s, e, label in entries:
                    uniq[(int(s), int(e))] = (int(s), int(e), str(label))

                entries = list(uniq.values())
                entries.sort(key=lambda t: (int(t[0]), int(t[1])))
                self._bubble_entries = entries

                if not entries:
                    self._bubble_entry_index = -1
                    try:
                        self._bubble_scroll_edit.setText("-")
                    except Exception:
                        pass
                    return

                # Clamp index.
                if self._bubble_entry_index < 0 or self._bubble_entry_index >= len(entries):
                    self._bubble_entry_index = 0
                try:
                    self._bubble_scroll_edit.setText(str(entries[self._bubble_entry_index][2]))
                except Exception:
                    pass

            def _set_scroller_to(n: int) -> bool:
                """Set the Bubble scroller display to the bubble item containing n."""
                try:
                    nn = int(n)
                except Exception:
                    return False
                if nn <= 0:
                    return False

                # Refresh entries from the PDF state.
                try:
                    _refresh_bubble_entries()
                except Exception:
                    pass

                entries = list(getattr(self, "_bubble_entries", []) or [])
                if not entries:
                    return False

                idx = None
                for i, (s, e, _label) in enumerate(entries):
                    try:
                        s = int(s)
                        e = int(e)
                    except Exception:
                        continue
                    if e < s:
                        e = s
                    if s <= nn <= e:
                        idx = int(i)
                        break

                if idx is None:
                    return False

                try:
                    self._bubble_entry_index = int(idx)
                except Exception:
                    self._bubble_entry_index = int(idx)

                try:
                    self._bubble_scroll_edit.setText(str(entries[int(idx)][2]))
                except Exception:
                    pass
                return True

            # Expose setter for the Find Bubble handler.
            try:
                self._bubble_scroller_set_to_number = _set_scroller_to
            except Exception:
                pass

            def _emit_and_select_current(center: bool = True) -> None:
                try:
                    idx = int(self._bubble_entry_index)
                except Exception:
                    idx = -1
                entries = list(getattr(self, "_bubble_entries", []) or [])
                if idx < 0 or idx >= len(entries):
                    return
                s, e, label = entries[idx]
                try:
                    self._bubble_scroll_edit.setText(str(label))
                except Exception:
                    pass
                try:
                    if hasattr(self._pdf_viewer, "select_bubble_number"):
                        # Selecting by start selects the whole range bubble item.
                        self._pdf_viewer.select_bubble_number(int(s), center=bool(center))
                except Exception:
                    pass
                try:
                    self.bubbleScrollSelected.emit(int(s), int(e))
                except Exception:
                    pass

            def _step(delta: int) -> None:
                entries = list(getattr(self, "_bubble_entries", []) or [])
                if not entries:
                    return
                try:
                    self._bubble_entry_index = max(0, min(len(entries) - 1, int(self._bubble_entry_index) + int(delta)))
                except Exception:
                    self._bubble_entry_index = 0
                _emit_and_select_current(center=True)

            self._bubble_scroll_up.clicked.connect(lambda _c=False: _step(-1))
            self._bubble_scroll_down.clicked.connect(lambda _c=False: _step(1))

            try:
                self._pdf_viewer.bubbles_changed.connect(lambda *_a, **_k: _refresh_bubble_entries())
            except Exception:
                pass

            _refresh_bubble_entries()
        except Exception:
            pass

        # Optional docking controls (for when this window is used as a tab).
        try:
            spacer = QWidget()
            spacer.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
            tb.addWidget(spacer)

            if self._pop_out_callback is not None:
                self._pop_btn = QPushButton("Pop Out")
                self._pop_btn.setToolTip("Undock this viewer into a separate window")
                self._pop_btn.clicked.connect(self._pop_out_callback)
                tb.addWidget(self._pop_btn)

            if self._dock_back_callback is not None:
                self._dock_btn = QPushButton("Dock Back")
                self._dock_btn.setToolTip("Dock this viewer back into the main tabs")
                self._dock_btn.clicked.connect(self._dock_back_callback)
                tb.addWidget(self._dock_btn)

            # Default: assume "docked" when created with a pop-out callback.
            self.set_docked_state(self._pop_out_callback is not None)
        except Exception:
            pass


    def set_docked_state(self, is_docked: bool) -> None:
        """Controls which docking buttons are visible.

        - Docked (in tab): show Pop Out, hide Dock Back
        - Undocked (separate window): hide Pop Out, show Dock Back
        """
        try:
            if self._pop_btn is not None:
                self._pop_btn.setVisible(bool(is_docked))
        except Exception:
            pass
        try:
            if self._dock_btn is not None:
                self._dock_btn.setVisible(not bool(is_docked))
        except Exception:
            pass

    def _build_docks(self) -> None:
        v = self._pdf_viewer

        # Properties dock
        props = QWidget()
        props_root = QVBoxLayout(props)
        props_root.setContentsMargins(8, 8, 8, 8)
        props_root.setSpacing(10)

        # ---- Bubble Properties ----
        bubble_group = QGroupBox("Bubble Properties")
        bubble_layout = QFormLayout(bubble_group)
        bubble_layout.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)

        size_row = QWidget()
        size_row_l = QHBoxLayout(size_row)
        size_row_l.setContentsMargins(0, 0, 0, 0)
        size_row_l.addWidget(v.size_slider)
        size_row_l.addWidget(v.size_label)

        line_row = QWidget()
        line_row_l = QHBoxLayout(line_row)
        line_row_l.setContentsMargins(0, 0, 0, 0)
        line_row_l.addWidget(v.line_slider)
        line_row_l.addWidget(v.line_label)

        bubble_layout.addRow("Shape", v.shape_combo)
        bubble_layout.addRow("Size", size_row)
        bubble_layout.addRow("Line", line_row)

        try:
            self._apply_theme_aware_combobox_style(v.shape_combo)
        except Exception:
            pass


        # Bubble/annotation color
        try:
            color_row = QWidget()
            color_row_l = QHBoxLayout(color_row)
            color_row_l.setContentsMargins(0, 0, 0, 0)
            color_row_l.setSpacing(6)

            self._bubble_color_swatch = QLabel("")
            try:
                self._bubble_color_swatch.setFixedSize(22, 22)
            except Exception:
                pass

            def _refresh_swatch() -> None:
                try:
                    qc = getattr(v, "bubble_color", None)
                    if qc is None or not qc.isValid():
                        self._bubble_color_swatch.setStyleSheet("")
                        return
                    self._bubble_color_swatch.setStyleSheet(f"background-color: {qc.name()};")
                except Exception:
                    pass

            _refresh_swatch()

            pick_btn = QPushButton("Choose…")
            pick_btn.setToolTip("Change bubble/annotation color")
            pick_btn.clicked.connect(lambda: self._pick_bubble_color())

            color_row_l.addWidget(self._bubble_color_swatch)
            color_row_l.addWidget(pick_btn)
            color_row_l.addStretch(1)
            bubble_layout.addRow("Color", color_row)
            self._refresh_bubble_color_swatch = _refresh_swatch
        except Exception:
            pass


        props_root.addWidget(bubble_group)

        # ---- Drawing Grid ----
        try:
            grid_group = QGroupBox("Drawing Grid")
            grid_layout = QFormLayout(grid_group)
            grid_layout.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)

            show_grid_cb = QCheckBox("Show grid")
            try:
                show_grid_cb.setChecked(bool(getattr(v, "grid_enabled", False)))
            except Exception:
                show_grid_cb.setChecked(False)
            grid_layout.addRow(show_grid_cb)

            left_spin = QDoubleSpinBox()
            top_spin = QDoubleSpinBox()
            width_spin = QDoubleSpinBox()
            height_spin = QDoubleSpinBox()
            for sp in (left_spin, top_spin, width_spin, height_spin):
                sp.setRange(0.0, 100.0)
                sp.setDecimals(1)
                sp.setSingleStep(0.5)
                sp.setSuffix("%")

            _grid_sync = {"busy": False}

            try:
                left_spin.setValue(float(getattr(v, "grid_left_pct", 0.0) or 0.0))
                top_spin.setValue(float(getattr(v, "grid_top_pct", 0.0) or 0.0))
                width_spin.setValue(float(getattr(v, "grid_width_pct", 100.0) or 100.0))
                height_spin.setValue(float(getattr(v, "grid_height_pct", 100.0) or 100.0))
            except Exception:
                pass

            grid_layout.addRow("Left", left_spin)
            grid_layout.addRow("Top", top_spin)
            grid_layout.addRow("Width", width_spin)
            grid_layout.addRow("Height", height_spin)

            def _apply_grid() -> None:
                try:
                    if _grid_sync.get("busy"):
                        return
                    if hasattr(v, "set_grid_bounds_pct"):
                        v.set_grid_bounds_pct(
                            float(left_spin.value()),
                            float(top_spin.value()),
                            float(width_spin.value()),
                            float(height_spin.value()),
                        )
                except Exception:
                    pass

            def _on_grid_bounds_changed(left_pct: float, top_pct: float, width_pct: float, height_pct: float) -> None:
                # Viewer dragged bounds -> update spinboxes without feeding back.
                _grid_sync["busy"] = True
                try:
                    try:
                        from PySide6.QtCore import QSignalBlocker
                        blockers = [QSignalBlocker(left_spin), QSignalBlocker(top_spin), QSignalBlocker(width_spin), QSignalBlocker(height_spin)]
                    except Exception:
                        blockers = []

                    left_spin.setValue(float(left_pct))
                    top_spin.setValue(float(top_pct))
                    width_spin.setValue(float(width_pct))
                    height_spin.setValue(float(height_pct))

                    # Ensure blockers are kept alive until after setValue calls.
                    _ = blockers
                except Exception:
                    pass
                finally:
                    _grid_sync["busy"] = False

            def _toggle_grid(on: bool) -> None:
                try:
                    if hasattr(v, "set_grid_enabled"):
                        v.set_grid_enabled(bool(on))
                except Exception:
                    pass

            show_grid_cb.toggled.connect(_toggle_grid)
            left_spin.valueChanged.connect(lambda _v: _apply_grid())
            top_spin.valueChanged.connect(lambda _v: _apply_grid())
            width_spin.valueChanged.connect(lambda _v: _apply_grid())
            height_spin.valueChanged.connect(lambda _v: _apply_grid())

            # Live-sync controls when user drags the bounds rectangle in the viewer.
            try:
                if hasattr(v, "grid_bounds_changed"):
                    v.grid_bounds_changed.connect(_on_grid_bounds_changed)
            except Exception:
                pass

            props_root.addWidget(grid_group)
        except Exception:
            pass

        # ---- Form 3 / Reference ----
        try:
            ref_group = QGroupBox("Reference")
            ref_layout = QFormLayout(ref_group)
            ref_layout.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)

            ref_combo = QComboBox()
            ref_combo.addItems(["Sheet and Zone", "Page Number", "None"])

            def _current_mode() -> str:
                try:
                    return str(getattr(v, "reference_location_mode", "sheet_zone") or "sheet_zone").lower()
                except Exception:
                    return "sheet_zone"

            mode = _current_mode()
            if mode in ("page", "page_label", "page_number", "page number"):
                ref_combo.setCurrentText("Page Number")
            elif mode in ("none", "off", "disable", "disabled"):
                ref_combo.setCurrentText("None")
            else:
                ref_combo.setCurrentText("Sheet and Zone")

            def _apply_mode(text: str) -> None:
                t = str(text or "").strip().lower()
                if t.startswith("sheet"):
                    v.reference_location_mode = "sheet_zone"
                elif t.startswith("page"):
                    v.reference_location_mode = "page_label"
                else:
                    v.reference_location_mode = "none"

            def _on_mode_changed(text: str) -> None:
                _apply_mode(text)
                try:
                    if hasattr(self._pdf_viewer, "bubbles_changed"):
                        self._pdf_viewer.bubbles_changed.emit(self._pdf_viewer.get_bubbled_numbers())
                except Exception:
                    pass

            ref_combo.currentTextChanged.connect(_on_mode_changed)
            ref_layout.addRow("Reference Location", ref_combo)

            try:
                self._apply_theme_aware_combobox_style(ref_combo)
            except Exception:
                pass

            props_root.addWidget(ref_group)
        except Exception:
            pass

        # ---- Enhance Image ----
        # Always visible/expanded (no checkbox) per request.
        try:
            enhance_group = QGroupBox("Enhance Image")
            enhance_group_l = QVBoxLayout(enhance_group)
            enhance_group_l.setContentsMargins(8, 8, 8, 8)
            enhance_group_l.setSpacing(6)

            # Ensure enhancement mode is enabled so sliders take effect.
            try:
                if hasattr(v, "set_enhance_mode"):
                    v.set_enhance_mode(True)
                else:
                    v.enhance_mode = True
            except Exception:
                pass

            enhance_group_l.addWidget(v.enhance_panel)
            try:
                v.enhance_panel.setVisible(True)
            except Exception:
                pass

            props_root.addWidget(enhance_group)
        except Exception:
            pass


        # ---- Existing Bubbles ----
        existing_group = None
        existing_layout = None
        try:
            existing_group = QGroupBox("Existing Bubbles")
            existing_layout = QFormLayout(existing_group)
            existing_layout.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)
        except Exception:
            existing_group = None
            existing_layout = None

        # Toggle auto-import of external annotations (overlay)
        try:
            v.auto_import_btn = QPushButton("Auto Import Ext. Bubbles")
            v.auto_import_btn.setCheckable(True)
            v.auto_import_btn.setChecked(bool(getattr(v, "auto_import_annots", True)))
            v.auto_import_btn.setToolTip(
                "If checked, existing annotations (e.g. Kofax) will be imported as bubbles if no app bubbles exist."
            )
            v.auto_import_btn.toggled.connect(lambda on: setattr(v, "auto_import_annots", bool(on)))
            v.drawing_saved.connect(lambda _path: v.auto_import_btn.setChecked(False))
            if existing_layout is not None:
                existing_layout.addRow("Import", v.auto_import_btn)
        except Exception:
            pass

        # Toggle visibility of existing/app bubbles
        try:
            v.show_external_btn = QPushButton("Ext. Bubbles (0)")
            v.show_external_btn.setCheckable(True)
            v.show_external_btn.setChecked(True)
            v.show_external_btn.setToolTip("Show/hide existing annotations not created by this app")

            def _toggle_external(on):
                v.show_external_annots = bool(on)
                v.render_current_page(target_scale=v.current_render_scale, source_rotation=v._last_render_rotation)

            v.show_external_btn.toggled.connect(_toggle_external)
            if existing_layout is not None:
                existing_layout.addRow("External", v.show_external_btn)

            v.show_internal_btn = QPushButton("App Bubbles (0)")
            v.show_internal_btn.setCheckable(True)
            v.show_internal_btn.setChecked(True)
            v.show_internal_btn.setToolTip("Show/hide bubbles created by this app")

            def _toggle_internal(on):
                v.show_internal_annots = bool(on)
                v.render_current_page(target_scale=v.current_render_scale, source_rotation=v._last_render_rotation)

            v.show_internal_btn.toggled.connect(_toggle_internal)
            if existing_layout is not None:
                existing_layout.addRow("App", v.show_internal_btn)

            def _update_counts(_bubbles=None):
                try:
                    page_internal, page_external = v.get_annotation_counts(v.current_page)
                    total_internal, total_external = v.get_total_annotation_counts()
                    v.show_internal_btn.setText(f"App Bubbles ({page_internal})({total_internal})")
                    v.show_external_btn.setText(f"Ext. Bubbles ({page_external})({total_external})")
                    status_text = (
                        f"Status: Showing {page_internal} App Bubbles (Total {total_internal}), "
                        f"{page_external} Existing Bubbles (Total {total_external})"
                    )
                    if hasattr(self, "status_label"):
                        self.status_label.setText(status_text)
                except Exception:
                    pass

            v.bubbles_changed.connect(_update_counts)
            try:
                _update_counts()
            except Exception:
                pass
        except Exception:
            pass


        # Delete embedded PDF annotations (aka existing bubbles)
        try:
            v.delete_pdf_annots_btn = QPushButton("Delete Ex. Bubbles")
            v.delete_pdf_annots_btn.setToolTip("Delete existing bubbles already embedded in the PDF")
            v.delete_pdf_annots_btn.clicked.connect(self._delete_pdf_annots_current_page)
            if existing_layout is not None:
                existing_layout.addRow("Delete", v.delete_pdf_annots_btn)
        except Exception:
            pass

        try:
            if existing_group is not None:
                props_root.addWidget(existing_group)
        except Exception:
            pass

        try:
            props_root.addStretch(1)
        except Exception:
            pass

        props_dock = QDockWidget("Properties", self)
        props_dock.setWidget(props)
        props_dock.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        self.addDockWidget(Qt.RightDockWidgetArea, props_dock)

        # Notes dock
        notes_page = QWidget()
        notes_l = QVBoxLayout(notes_page)
        notes_l.setContentsMargins(6, 6, 6, 6)
        notes_l.setSpacing(6)
        notes_l.addWidget(v.add_note_region_btn)
        notes_l.addWidget(v.clear_note_regions_btn)

        mode_row = QWidget()
        mode_row_l = QHBoxLayout(mode_row)
        mode_row_l.setContentsMargins(0, 0, 0, 0)
        mode_row_l.addWidget(QLabel("Mode:"))
        mode_row_l.addWidget(v.notes_mode_combo)
        try:
            self._apply_theme_aware_combobox_style(v.notes_mode_combo)
        except Exception:
            pass
        notes_l.addWidget(mode_row)

        notes_l.addWidget(v.extract_notes_btn)
        notes_l.addStretch(1)

        # Actions dock
        actions_page = QWidget()
        actions_l = QVBoxLayout(actions_page)
        actions_l.setContentsMargins(6, 6, 6, 6)
        actions_l.setSpacing(6)
        try:
            actions_l.addWidget(v.clear_btn)
            actions_l.addWidget(v.select_page_btn)
            actions_l.addWidget(v.select_all_btn)
            actions_l.addWidget(v.copy_btn)
            actions_l.addWidget(v.paste_btn)
        except Exception:
            pass
        actions_l.addStretch(1)

        notes_dock = QDockWidget("Notes", self)
        notes_dock.setWidget(notes_page)
        notes_dock.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        self.addDockWidget(Qt.RightDockWidgetArea, notes_dock)

        actions_dock = QDockWidget("Actions", self)
        actions_dock.setWidget(actions_page)
        actions_dock.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        self.addDockWidget(Qt.RightDockWidgetArea, actions_dock)

        self.tabifyDockWidget(props_dock, notes_dock)
        self.tabifyDockWidget(props_dock, actions_dock)
        notes_dock.raise_()

        # Enhance dock removed (Enhance moved into Properties).

    def _pick_bubble_color(self) -> None:
        v = self._pdf_viewer
        try:
            cur = getattr(v, "bubble_color", None)
            if cur is None:
                cur = None
        except Exception:
            cur = None
        try:
            color = QColorDialog.getColor(cur, self, "Select Bubble Color")
        except Exception:
            return
        try:
            if not color.isValid():
                return
        except Exception:
            pass
        try:
            v.set_bubble_color(color)
        except Exception:
            pass
        try:
            if hasattr(self, "_refresh_bubble_color_swatch"):
                self._refresh_bubble_color_swatch()
        except Exception:
            pass

    def _delete_pdf_annots_current_page(self) -> None:
        v = self._pdf_viewer
        all_pages = False
        try:
            from PySide6.QtWidgets import QMessageBox

            if int(getattr(v, "total_pages", 0) or 0) > 1:
                mb = QMessageBox(self)
                mb.setIcon(QMessageBox.Warning)
                mb.setWindowTitle("Delete Existing Bubbles")
                mb.setText("Delete existing bubbles already embedded in the PDF?\n\nThis cannot be undone.")
                cur_btn = mb.addButton("Current Page", QMessageBox.AcceptRole)
                all_btn = mb.addButton("All Pages", QMessageBox.DestructiveRole)
                mb.addButton(QMessageBox.Cancel)
                mb.setDefaultButton(QMessageBox.Cancel)
                mb.exec()
                clicked = mb.clickedButton()
                if clicked is None or clicked == mb.button(QMessageBox.Cancel):
                    return
                all_pages = bool(clicked == all_btn)
            else:
                mb = QMessageBox(self)
                mb.setIcon(QMessageBox.Warning)
                mb.setWindowTitle("Delete Existing Bubbles")
                mb.setText("Delete existing bubbles already embedded in the PDF on the current page?\n\nThis cannot be undone.")
                mb.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
                mb.setDefaultButton(QMessageBox.No)
                if mb.exec() != QMessageBox.Yes:
                    return
        except Exception:
            # Default to current page.
            all_pages = False

        deleted = 0
        try:
            deleted = int(v.delete_pdf_annotations(all_pages=bool(all_pages)) or 0)
        except Exception:
            deleted = 0

        try:
            from PySide6.QtWidgets import QMessageBox

            if all_pages:
                QMessageBox.information(self, "Existing Bubbles", f"Deleted {deleted} bubble annotation(s) across all pages.")
            else:
                QMessageBox.information(self, "Existing Bubbles", f"Deleted {deleted} bubble annotation(s) on this page.")
        except Exception:
            pass

    def _build_status_bar(self) -> None:
        v = self._pdf_viewer
        sb = self.statusBar()

        # Status label for annotation counts
        self.status_label = QLabel("Status: Ready")
        sb.addWidget(self.status_label, 1)

        w = QWidget()
        l = QHBoxLayout(w)
        l.setContentsMargins(6, 0, 6, 0)
        l.setSpacing(6)

        l.addWidget(QLabel("Page:"))
        l.addWidget(v.page_combo)
        l.addWidget(v.prev_btn)
        l.addWidget(v.next_btn)

        l.addWidget(QLabel("| Zoom:"))
        l.addWidget(v.zoom_out_btn)
        l.addWidget(v.zoom_fit_btn)
        l.addWidget(v.zoom_100_btn)
        l.addWidget(v.zoom_in_btn)
        l.addWidget(v.zoom_label)

        l.addWidget(QLabel("| Rotate:"))
        l.addWidget(v.rotate_left_btn)
        l.addWidget(v.rotate_right_btn)

        sb.addPermanentWidget(w, 1)
