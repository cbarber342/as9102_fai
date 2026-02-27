"""Core logic for the AS9102 FAI project.
"""

from __future__ import annotations
import os
import sys
import logging
from PySide6.QtWidgets import QApplication
from as9102_fai.gui.main_window import MainWindow
from as9102_fai.logging_utils import configure_logging

def run() -> None:
    """Launch the AS9102 FAI GUI Application."""
    debug_pdf = str(os.environ.get("AS9102_DEBUG_PDF", "")).strip().lower() in ("1", "true", "yes", "on")
    debug_gdt = str(os.environ.get("AS9102_DEBUG_GDT", "")).strip().lower() in ("1", "true", "yes", "on")

    # Console logging is useful for both PDF and GD&T debugging.
    console_debug = bool(debug_pdf or debug_gdt)
    configure_logging(debug=console_debug)

    log = logging.getLogger(__name__)
    log.debug(
        "Starting AS9102 FAI GUI (console_debug=%s debug_pdf=%s debug_gdt=%s)",
        console_debug,
        debug_pdf,
        debug_gdt,
    )
    if debug_gdt:
        log.debug("Env AS9102_DEBUG_GDT=%r", os.environ.get("AS9102_DEBUG_GDT"))

    # Check if QApplication already exists
    app = QApplication.instance()
    if not app:
        app = QApplication(sys.argv)
    
    window = MainWindow()
    window.showMaximized()
    
    sys.exit(app.exec())
