"""
web_automation.py
=================
Modulo responsable de automatizar la sesion web en el portal Recave:
  1. Login con credenciales.
  2. Navegacion hacia la seccion de carga de cartera.
  3. Seleccion de producto y delimitador.
  4. Subida del archivo CSV via Selenium.
  5. Espera del procesamiento y cierre de sesion.
"""
import time
from pathlib import Path
from typing import Optional

import pyautogui
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By

from config import USUARIO_WEB, PASSWORD_WEB, URL_DESTINO
from exceptions import WebAutomationError
from logger import get_logger

log = get_logger(__name__)

# ── Constantes ─────────────────────────────────────────────────────────────────
CHROME_VERSION: int         = 147     # Version de Chrome instalada en el equipo
IMAGE_SEARCH_TIMEOUT: int   = 15      # Segundos maximos buscando una imagen
IMAGE_CONFIDENCE: float     = 0.8     # Nivel de confianza para pyautogui
MOUSE_MOVE_DURATION: float  = 0.6     # Segundos que tarda el mouse en moverse
HUMAN_REACTION_DELAY: float = 0.5     # Pausa de "reaccion humana" antes del clic
POST_ACTION_DELAY: float    = 1.5     # Pausa despues de cada accion de pyautogui
UPLOAD_WAIT_SECONDS: int    = 10      # Segundos a esperar tras inyectar el archivo
PROCESSING_WAIT_SECONDS: int = 120    # Segundos a esperar el procesamiento del servidor


# ── Helpers privados ───────────────────────────────────────────────────────────

def _wait_and_click(
    image_path: str,
    timeout: int = IMAGE_SEARCH_TIMEOUT,
    confidence: float = IMAGE_CONFIDENCE,
    move_only: bool = False,
) -> None:
    """
    Busca una imagen en pantalla de forma repetida hasta encontrarla o agotar
    el tiempo de espera, luego hace clic (o solo mueve el cursor si move_only=True).

    Args:
        image_path: Ruta a la imagen de referencia (.png).
        timeout:    Tiempo maximo de busqueda en segundos.
        confidence: Nivel de confianza de la coincidencia (0.0 - 1.0).
        move_only:  Si es True, solo mueve el cursor sin hacer clic.

    Raises:
        WebAutomationError: Si la imagen no se encuentra en el tiempo dado.
    """
    log.debug("Buscando imagen: %s", image_path)
    deadline = time.monotonic() + timeout

    while time.monotonic() < deadline:
        try:
            coords = pyautogui.locateCenterOnScreen(image_path, confidence=confidence)
            if coords is not None:
                x, y = coords[0], coords[1]
                time.sleep(HUMAN_REACTION_DELAY)

                if move_only:
                    pyautogui.moveTo(x, y, duration=MOUSE_MOVE_DURATION)
                else:
                    pyautogui.moveTo(x, y, duration=MOUSE_MOVE_DURATION)
                    time.sleep(0.2)
                    pyautogui.click(x, y)

                time.sleep(POST_ACTION_DELAY)
                return
        except Exception:
            pass
        time.sleep(0.5)

    raise WebAutomationError(
        f"No se encontro la imagen '{image_path}' en pantalla "
        f"despues de {timeout} segundos."
    )


def _inject_file(driver: uc.Chrome, file_path: Path) -> None:
    """
    Inyecta el archivo al input de tipo file de la pagina via Selenium,
    evitando el dialogo nativo del sistema operativo.

    Args:
        driver:    Instancia del WebDriver activo.
        file_path: Ruta absoluta al archivo a subir.

    Raises:
        WebAutomationError: Si el elemento de subida no se encuentra.
    """
    log.info("Inyectando archivo via Selenium: %s", file_path.name)
    try:
        input_element = driver.find_element(
            By.ID, "ctl00_CPHMaster_AsyncFileUpload1_ctl02"
        )
        input_element.send_keys(str(file_path))
        log.debug("Archivo inyectado correctamente. Esperando procesamiento en web...")
        time.sleep(UPLOAD_WAIT_SECONDS)
    except Exception as exc:
        raise WebAutomationError(f"No se pudo inyectar el archivo: {exc}") from exc


# ── Funcion publica ────────────────────────────────────────────────────────────

def run_web_automation(archivo_a_subir: str) -> None:
    """
    Orquesta la sesion completa de automatizacion web:
    login -> navegacion -> seleccion -> subida -> espera -> logout.

    Args:
        archivo_a_subir: Ruta absoluta al archivo CSV generado por excel_processor.

    Raises:
        WebAutomationError: Si el navegador no puede iniciarse.
    """
    file_path = Path(archivo_a_subir).resolve()
    log.info("Iniciando navegador web (Chrome %d)...", CHROME_VERSION)

    options = uc.ChromeOptions()
    options.add_experimental_option("prefs", {
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
    })
    # Descomenta la siguiente linea para correr en modo invisible (headless):
    # options.add_argument("--headless")

    try:
        driver = uc.Chrome(options=options, version_main=CHROME_VERSION)
    except Exception as exc:
        raise WebAutomationError(f"No se pudo abrir el navegador: {exc}") from exc

    try:
        log.info("Navegando a: %s", URL_DESTINO)
        driver.get(URL_DESTINO)

        # ── 1. Login ───────────────────────────────────────────────────────────
        log.info("[1/5] Iniciando sesion...")
        _wait_and_click("images/inicio_sesion.png")
        pyautogui.write(USUARIO_WEB, interval=0.1)
        pyautogui.press("tab")
        time.sleep(0.5)
        pyautogui.write(PASSWORD_WEB, interval=0.1)
        pyautogui.press("enter")

        # ── 2. Navegar a Cartera ───────────────────────────────────────────────
        log.info("[2/5] Navegando a Cartera...")
        _wait_and_click("images/apartado.png", move_only=True)
        _wait_and_click("images/cartera.png")

        # ── 3. Seleccionar Producto ────────────────────────────────────────────
        log.info("[3/5] Seleccionando producto...")
        _wait_and_click("images/producto.png")
        _wait_and_click("images/seleccion_producto.png")

        # ── 4. Subir archivo ───────────────────────────────────────────────────
        log.info("[4/5] Subiendo archivo...")
        time.sleep(2)  # Pausa tecnica para que el DOM este listo
        _inject_file(driver, file_path)

        # ── 5. Seleccionar delimitador ─────────────────────────────────────────
        log.info("[5/5] Seleccionando delimitador (coma)...")
        _wait_and_click("images/delimitador.png")
        _wait_and_click("images/coma.png")

        # ── 6. Cargar y esperar ────────────────────────────────────────────────
        log.info("Presionando Cargar y esperando procesamiento del servidor...")
        _wait_and_click("images/cargar.png")
        log.info(
            "Carga iniciada. Esperando %d segundos por el servidor...",
            PROCESSING_WAIT_SECONDS,
        )
        time.sleep(PROCESSING_WAIT_SECONDS)

        # ── 7. Logout ──────────────────────────────────────────────────────────
        log.info("Cerrando sesion...")
        _wait_and_click("images/cerrar_sesion.png")

        log.info("Automatizacion completada exitosamente.")

    except WebAutomationError:
        raise
    except Exception as exc:
        raise WebAutomationError(f"Error inesperado en la automatizacion: {exc}") from exc
    finally:
        input("\nPresiona ENTER para cerrar el navegador... ")
        try:
            driver.quit()
            log.debug("Navegador cerrado correctamente.")
        except Exception:
            pass
