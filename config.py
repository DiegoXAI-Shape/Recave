import os
from dotenv import load_dotenv

# Cargar las variables de entorno desde el archivo .env
load_dotenv()

# ==============================================================================
# CONFIGURACIONES GENERALES
# ==============================================================================

# 1. Contraseña del archivo Excel encriptado (viene de tu archivo .env)
EXCEL_PASSWORD = os.getenv("excel_password")
USUARIO_WEB = os.getenv("user")
PASSWORD_WEB = os.getenv("pwd")

# 2. Configuración de archivos
EXCEL_ORIGIN = "NOMBRE_DEL_ARCHIVO.xlsx"  # Nombre del archivo original
EXCEL_SHEET = "Nombre de la Hoja"                     # Nombre de la pestaña a leer
EXCEL_CLEAN = "resultado_limpio.xlsx"                 # Nombre del archivo final a guardar

# 3. Configuración Web
URL_DESTINO = "https://TU_PORTAL_WEB/ruta/LogIn.aspx"                # URL de la página web a automatizar
