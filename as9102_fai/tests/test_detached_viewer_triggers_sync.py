def test_detached_viewer_bubbles_changed_triggers_full_sync() -> None:
    """Regression: detached Drawing Viewer emits bubbles_changed.

    That signal must trigger the same pipeline as embedded viewer sync so that
    thread-derived row auto-remapping can run (bubble 5 -> 7, etc.).
    """

    pyside = __import__("importlib").import_module("importlib").util.find_spec("PySide6")
    if pyside is None:
        return

    from as9102_fai.gui.main_window import MainWindow

    calls: list[set[int]] = []

    class Dummy:
        _last_bubbled_numbers = None

        def _sync_bubbles_to_form3(self, bubbled_numbers=None):
            calls.append(set(bubbled_numbers or set()))

        def _update_form3_bubble_fills(self, bubbled_numbers):
            raise AssertionError("Should not use legacy path")

        _form_viewers = {}

    d = Dummy()

    MainWindow._on_drawing_bubbles_changed(d, {"5", 6})

    assert calls == [{5, 6}]
