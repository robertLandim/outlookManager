import os
import logging
from logging.handlers import RotatingFileHandler

# Certifique-se de que a pasta "logs" existe
LOG_DIR = "logs"
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)

LOG_FILE = os.path.join(LOG_DIR, "outlook_manager.log")

# Configuração do handler com rotação, encoding utf-8
handler = RotatingFileHandler(
    LOG_FILE, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
)

formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(filename)s - %(message)s')
handler.setFormatter(formatter)

logger = logging.getLogger("outlook_manager")
logger.setLevel(logging.INFO)
logger.addHandler(handler)
logger.propagate = False  # Para evitar logs duplicados ao importar em outros módulos