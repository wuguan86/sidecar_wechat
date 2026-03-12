import logging
import logging.handlers
import os
from .config import BridgeConfig

def setup_logging(cfg: BridgeConfig) -> logging.Logger:
    logger = logging.getLogger("wechat_bridge")
    # Set log level
    logger.setLevel(getattr(logging, cfg.log_level.upper(), logging.INFO))

    if logger.handlers:
        return logger

    # Format: Time [Level] [Thread] Message
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] [%(threadName)s] %(message)s")

    os.makedirs(os.path.dirname(cfg.log_file), exist_ok=True)
    
    # Use RotatingFileHandler
    file_handler = logging.handlers.RotatingFileHandler(
        cfg.log_file, maxBytes=cfg.log_max_bytes, backupCount=cfg.log_backup_count, encoding="utf-8"
    )
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # Console output
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    return logger
