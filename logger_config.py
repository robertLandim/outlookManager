import logging
import os
from logging.handlers import RotatingFileHandler

from config import APP_ROOT

LOG_DIR = APP_ROOT / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)

LOG_FILE = LOG_DIR / "outlook_manager.log"

handler = RotatingFileHandler(
    str(LOG_FILE), maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
)

formatter = logging.Formatter(
    "%(asctime)s - %(levelname)s - %(filename)s - %(message)s"
)
handler.setFormatter(formatter)

logger = logging.getLogger("outlook_manager")
logger.setLevel(logging.INFO)
logger.addHandler(handler)
logger.propagate = False

if os.environ.get("DEBUG", "").strip().lower() in ("1", "true", "yes"):
    _console = logging.StreamHandler()
    _console.setFormatter(formatter)
    _console.setLevel(logging.DEBUG)
    logger.addHandler(_console)
    logger.setLevel(logging.DEBUG)
