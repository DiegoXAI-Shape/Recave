"""
main.py
=======
Punto de entrada del proyecto Recave.

Flujo:
  1. Procesa el archivo Excel → genera CSV limpio.
  2. Automatiza la sesion web para cargar el CSV.
"""
import sys
from datetime import datetime

from exceptions import RecaveError, ConfigError
from logger import get_logger
import config
from excel_processor import process_excel
from web_automation import run_web_automation

log = get_logger(__name__)

# ── Banner ─────────────────────────────────────────────────────────────────────
_BANNER = """
╔══════════════════════════════════════════════════╗
║           RECAVE — Automatizacion de Carga        ║
║         {fecha:<42} ║
╚══════════════════════════════════════════════════╝
"""


def main() -> None:
    """Orquesta el pipeline completo: procesamiento Excel + automatizacion web."""
    print(_BANNER.format(fecha=datetime.now().strftime("%Y-%m-%d  %H:%M")))

    # ── Paso 1: Procesar Excel ─────────────────────────────────────────────────
    log.info("=" * 55)
    log.info("PASO 1 — Procesando Excel")
    log.info("=" * 55)
    csv_file = process_excel(
        password=config.EXCEL_PASSWORD,
        origin_file=config.EXCEL_ORIGIN,
        sheet_name=config.EXCEL_SHEET,
    )

    # ── Paso 2: Automatizacion Web ─────────────────────────────────────────────
    log.info("=" * 55)
    log.info("PASO 2 — Iniciando automatizacion web")
    log.info("=" * 55)
    run_web_automation(str(csv_file))


if __name__ == "__main__":
    try:
        main()
    except ConfigError as exc:
        log.error("[CONFIG] %s", exc)
        sys.exit(1)
    except RecaveError as exc:
        log.error("[ERROR] %s", exc)
        sys.exit(1)
    except KeyboardInterrupt:
        log.warning("Ejecucion interrumpida por el usuario.")
        sys.exit(0)
