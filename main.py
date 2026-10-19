"""Entry point: configures logging and launches the desktop app."""

from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

import config


def configure_logging() -> None:
    """Log to logs/app.log (if writable) and to the console."""
    handlers: list[logging.Handler] = [logging.StreamHandler()]
    try:
        log_dir = Path(__file__).resolve().parent / config.LOG_DIR_NAME
        log_dir.mkdir(exist_ok=True)
        handlers.append(
            RotatingFileHandler(
                log_dir / config.LOG_FILE_NAME, maxBytes=512_000, backupCount=2,
                encoding="utf-8",
            )
        )
    except OSError:
        pass  # console logging only
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        handlers=handlers,
    )


def main() -> None:
    configure_logging()
    from agent import CaseAgent
    from llm_agent import LLMController
    from ui import CaseAgentApp  # imported late so logging is ready first

    llm = LLMController.from_environment()  # None when no API key is set
    CaseAgentApp(CaseAgent(llm=llm)).mainloop()


if __name__ == "__main__":
    main()
