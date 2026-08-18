"""
logger.py
=========
Configuracion centralizada del logger para el proyecto Recave.

Uso:
    from logger import get_logger
    log = get_logger(__name__)
    log.info("Mensaje informativo")
    log.error("Algo salio mal")
"""
import logging
import sys
from datetime import datetime
from pathlib import Path

# ── Constantes ─────────────────────────────────────────────────────────────────
LOG_DIR = Path(__file__).parent / "logs"
LOG_FORMAT = "%(asctime)s  %(levelname)-8s  %(name)s — %(message)s"
DATE_FORMAT = "%Y-%m-%d %H:%M:%S"

# Codigos ANSI para colores en consola
_COLORS = {
    "DEBUG":    "\033[36m",   # Cyan
    "INFO":     "\033[32m",   # Verde
    "WARNING":  "\033[33m",   # Amarillo
    "ERROR":    "\033[31m",   # Rojo
    "CRITICAL": "\033[35m",   # Magenta
    "RESET":    "\033[0m",
}


class _ColorFormatter(logging.Formatter):
    """Formateador que colorea el nivel de log en la salida de consola."""

    def format(self, record: logging.LogRecord) -> str:
        color = _COLORS.get(record.levelname, _COLORS["RESET"])
        reset = _COLORS["RESET"]
        record.levelname = f"{color}{record.levelname}{reset}"
        return super().format(record)


def _setup_logger(name: str = "recave") -> logging.Logger:
    """
    Configura y devuelve el logger raiz del proyecto.

    - Consola: colores ANSI, nivel INFO.
    - Archivo: sin colores, nivel DEBUG, en logs/recave_YYYYMMDD.log.
    """
    logger = logging.getLogger(name)

    if logger.handlers:
        # Evita duplicar handlers si get_logger se llama varias veces
        return logger

    logger.setLevel(logging.DEBUG)

    # ── Handler de consola ──────────────────────────────────────────────────────
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(_ColorFormatter(LOG_FORMAT, datefmt=DATE_FORMAT))
    logger.addHandler(console_handler)

    # ── Handler de archivo ──────────────────────────────────────────────────────
    LOG_DIR.mkdir(exist_ok=True)
    log_filename = LOG_DIR / f"recave_{datetime.now().strftime('%Y%m%d')}.log"
    file_handler = logging.FileHandler(log_filename, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(logging.Formatter(LOG_FORMAT, datefmt=DATE_FORMAT))
    logger.addHandler(file_handler)

    return logger


def get_logger(name: str | None = None) -> logging.Logger:
    """
    Devuelve un logger hijo del logger raiz del proyecto.

    Args:
        name: Nombre del modulo (usa __name__ por convencion).

    Returns:
        logging.Logger: Logger configurado.
    """
    # Aseguramos que el logger raiz este configurado
    _setup_logger("recave")

    if name is None or name == "recave":
        return logging.getLogger("recave")

    return logging.getLogger(f"recave.{name}")
