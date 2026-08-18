"""
config.py
=========
Carga y valida toda la configuracion del proyecto Recave desde variables
de entorno (archivo .env). Si alguna variable critica no esta definida,
el programa falla inmediatamente con un mensaje claro.
"""
import os
from dotenv import load_dotenv
from exceptions import ConfigError

# Cargar las variables de entorno desde el archivo .env
load_dotenv()


def _require_env(key: str) -> str:
    """
    Obtiene una variable de entorno y lanza ConfigError si no existe.

    Args:
        key: Nombre de la variable de entorno.

    Returns:
        El valor de la variable.

    Raises:
        ConfigError: Si la variable no esta definida o esta vacia.
    """
    value = os.getenv(key)
    if not value:
        raise ConfigError(
            f"Falta la variable de entorno obligatoria: '{key}'. "
            "Asegurate de tener tu archivo .env configurado correctamente."
        )
    return value


# ==============================================================================
# CREDENCIALES (obligatorias — el programa falla si no estan en .env)
# ==============================================================================
EXCEL_PASSWORD: str = _require_env("excel_password")
USUARIO_WEB: str    = _require_env("user")
PASSWORD_WEB: str   = _require_env("pwd")

# ==============================================================================
# ARCHIVOS
# ==============================================================================
EXCEL_ORIGIN: str = "NOMBRE_DEL_ARCHIVO.xlsx"  # Archivo original
EXCEL_SHEET: str  = "Nombre de la Hoja"                      # Pestana a leer
EXCEL_CLEAN: str  = "resultado_limpio.xlsx"                  # Archivo de salida (ya no usado)

# ==============================================================================
# WEB
# ==============================================================================
URL_DESTINO: str = "https://TU_PORTAL_WEB/ruta/LogIn.aspx"
