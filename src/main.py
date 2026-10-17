"""Application entry point."""

from src.ui import FormatApp


def main() -> None:
    """Main application entry point."""
    app = FormatApp()
    app.mainloop()


if __name__ == "__main__":
    main()
