from __future__ import annotations

import logging
import os
from pathlib import Path
from logging.handlers import RotatingFileHandler


def configure_logging(*, debug: bool = False, log_path: str | None = None) -> None:
    """Configure app-wide logging.

    - Always logs to a rotating file under ./logs
    - Optionally logs to console when debug is enabled
    """

    level = logging.DEBUG if debug else logging.INFO

    if log_path is None:
        env_log_dir = os.environ.get("AS9102_FAI_LOG_DIR")
        if env_log_dir:
            log_dir = Path(env_log_dir)
            log_dir.mkdir(parents=True, exist_ok=True)
        else:
            # Prefer ./logs next to the current working directory for dev convenience,
            # but fall back to user-local app data if CWD isn't writable.
            cwd_logs = Path.cwd() / "logs"
            try:
                cwd_logs.mkdir(parents=True, exist_ok=True)
                log_dir = cwd_logs
            except Exception:
                appdata = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
                fallback = (Path(appdata) / "AS9102_FAI" / "logs") if appdata else cwd_logs
                fallback.mkdir(parents=True, exist_ok=True)
                log_dir = fallback

        log_path = str(log_dir / "as9102_fai_gui.log")

    root = logging.getLogger()
    root.setLevel(level)

    # Avoid duplicating handlers if run() is called more than once.
    if getattr(root, "_as9102_configured", False):
        return

    fmt = logging.Formatter(
        fmt="%(asctime)s %(levelname)s %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = RotatingFileHandler(log_path, maxBytes=1_000_000, backupCount=3, encoding="utf-8")
    file_handler.setLevel(level)
    file_handler.setFormatter(fmt)
    root.addHandler(file_handler)

    if debug:
        console = logging.StreamHandler()
        console.setLevel(level)
        console.setFormatter(fmt)
        root.addHandler(console)

    setattr(root, "_as9102_configured", True)
