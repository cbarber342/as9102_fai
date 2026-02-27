"""Entry point for the AS9102 FAI GUI application."""

from .core import run


def main() -> None:
    """Run the AS9102 FAI helper."""
    print("Starting AS9102 FAI application...")
    run()


if __name__ == "__main__":  # pragma: no cover
    main()
