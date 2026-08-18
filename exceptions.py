"""
exceptions.py
=============
Excepciones personalizadas para el proyecto Recave.
Permite un manejo de errores tipado y claro en todo el flujo.
"""


class RecaveError(Exception):
    """Excepcion base del proyecto Recave."""


class ConfigError(RecaveError):
    """Se lanza cuando la configuracion o el archivo .env es invalido o incompleto."""


class ExcelProcessError(RecaveError):
    """Se lanza cuando ocurre un error al desencriptar, leer o procesar el archivo Excel."""


class WebAutomationError(RecaveError):
    """Se lanza cuando falla algun paso de la automatizacion web (login, subida, etc.)."""
