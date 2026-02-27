from __future__ import annotations

import os
import shutil
from pathlib import Path
from typing import Optional


def find_tesseract_cmd(explicit: str | None = None) -> Optional[str]:
    """Find a usable Tesseract executable path.

    Search order:
    - explicit argument
    - env vars: AS9102_TESSERACT_CMD, TESSERACT_CMD, TESSERACT_PATH
    - common Windows install locations
    - PATH via shutil.which('tesseract')

    Returns a string path if found, else None.
    """

    candidates: list[str] = []

    if explicit:
        candidates.append(str(explicit))

    for env_name in ("AS9102_TESSERACT_CMD", "TESSERACT_CMD", "TESSERACT_PATH"):
        val = os.environ.get(env_name)
        if val:
            candidates.append(val)

    candidates.extend(
        [
            r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe",
            r"C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe",
            r"C:\\ProgramData\\chocolatey\\bin\\tesseract.exe",
        ]
    )

    which = shutil.which("tesseract")
    if which:
        candidates.append(which)

    for candidate in candidates:
        try:
            path = Path(candidate)
            if path.is_dir():
                path = path / "tesseract.exe"
            if path.exists():
                return str(path)
        except Exception:
            continue

    return None


def configure_pytesseract(explicit: str | None = None) -> bool:
    """Configure pytesseract to use a discovered tesseract executable.

    Returns True if it successfully configured a valid path, otherwise False.
    """

    cmd = find_tesseract_cmd(explicit)
    if not cmd:
        return False

    try:
        import pytesseract

        pytesseract.pytesseract.tesseract_cmd = cmd
        return True
    except Exception:
        return False
