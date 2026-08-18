"""
excel_processor.py
==================
Modulo responsable de:
  1. Desencriptar el archivo Excel protegido.
  2. Leer y mapear las columnas relevantes.
  3. Limpiar y normalizar los datos (telefonos, saldos).
  4. Exportar el resultado a un archivo CSV listo para cargar.
"""
import datetime
import os
import shutil
import tempfile
from pathlib import Path
from typing import Optional

import msoffcrypto
import pandas as pd

from exceptions import ExcelProcessError
from logger import get_logger

log = get_logger(__name__)

# ── Constantes ─────────────────────────────────────────────────────────────────
PHONE_COLUMNS = [
    "Telefono", "Tel Casa", "Tel Oficina",
    "Tel1", "Tel2", "Tel3", "Tel4", "Tel5", "Tel6", "Tel7", "Tel8",
]
MONEY_COLUMNS = [
    "Capital", "Responsa", "ADEUDO TOTAL", "Saldos Vdos", "VEN_CANOTO",
]


# ── Funciones privadas ─────────────────────────────────────────────────────────

def _decrypt_excel(origin_file: str, password: str) -> Path:
    """
    Desencripta un archivo Excel protegido con contrasena y lo guarda en un
    archivo temporal. Si el archivo ya no esta encriptado, lo copia directamente.

    Args:
        origin_file: Ruta al archivo Excel original.
        password:    Contrasena de desencriptacion.

    Returns:
        Path al archivo temporal desencriptado.

    Raises:
        ExcelProcessError: Si el archivo no existe.
    """
    origin_path = Path(origin_file)
    if not origin_path.exists():
        raise ExcelProcessError(f"No se encontro el archivo: '{origin_file}'")

    log.info("Desencriptando archivo: %s", origin_file)

    temp_fd, temp_path_str = tempfile.mkstemp(suffix=".xlsx")
    os.close(temp_fd)
    temp_path = Path(temp_path_str)

    try:
        with origin_path.open("rb") as file_in:
            office_file = msoffcrypto.OfficeFile(file_in)
            office_file.load_key(password=password)
            with temp_path.open("wb") as file_out:
                office_file.decrypt(file_out)
        log.debug("Desencriptacion exitosa.")
    except Exception:
        log.warning(
            "El archivo parece ya estar desencriptado o msoffcrypto no pudo leerlo. "
            "Usando archivo directo..."
        )
        shutil.copy(origin_path, temp_path)

    return temp_path


def _build_lookup(df_recave: pd.DataFrame) -> dict:
    """
    Construye el diccionario equivalente al BUSCARV de Excel:
    columna H (indice 7) → columna BE (indice 56).

    Args:
        df_recave: DataFrame de la hoja 'Nombre de la Hoja'.

    Returns:
        Diccionario {clave_h: valor_be}.
    """
    col_h = df_recave.columns[7]
    col_be = df_recave.columns[56]
    return dict(zip(df_recave[col_h], df_recave[col_be]))


def _build_dataframe(df_origen: pd.DataFrame, dic_vlookup: dict) -> pd.DataFrame:
    """
    Mapea las columnas del DataFrame de origen al formato de destino.

    Args:
        df_origen:    DataFrame con los datos de la hoja principal.
        dic_vlookup:  Diccionario de busqueda (BUSCARV).

    Returns:
        Nuevo DataFrame con las columnas en el formato requerido.
    """
    return pd.DataFrame({
        "Descrip Depto": df_origen.apply(
            lambda row: "Consumo B1"
            if row.get("Bucket") == "B1" and row.get("Descrip Depto") == "Consumo"
            else row.get("Descrip Depto"),
            axis=1,
        ),
        "Prod":                      df_origen["Prod"],
        "N Credito":                 df_origen["Nº Credito"],
        "Cta Che":                   df_origen["Cta Che"],
        "Cont":                      df_origen["Cont"],
        "N Pago Vdo":                df_origen["Nº Pago Vdo"],
        "Plzo":                      df_origen["Plzo"],
        "Nom Cliente":               df_origen["Nom Cliente"],
        "Fecha Vto":                 None,
        "Amo Dias":                  df_origen["Amo Dias"],
        "Dias Vdos":                 df_origen["Dias Vencidos"],
        "Gestor Nuevo":              None,
        "Saldos Vdos":               df_origen["SaldoPagarUltimoCorte"],
        "Responsa":                  df_origen["Responsa"],
        "Descri Credit":             df_origen["Descp Prodcuto"],
        "VEN_TIPCRE":                df_origen.get("VTO"),
        "Capital":                   df_origen["CAPITAL"],
        "Interes Nor":               None,
        "Int Mor":                   None,
        "Iva Fac":                   None,
        "Gasto Cob":                 None,
        "Iva Gtos":                  None,
        "Gastos Admvos":             None,
        "ADEUDO TOTAL":              None,
        "FINIQUITO CON DESCUENTO":   None,
        "Calle":                     df_origen["CALLE"],
        "N Calle":                   df_origen["NUM CALLE"],
        "Colonia":                   df_origen["Colonia"],
        "Codigo Postal":             df_origen["Codigo Postal"],
        "Municipio":                 df_origen["Municipio"],
        "Estado":                    df_origen["Estado"],
        "Tipo Auto":                 None,
        "FechaInicio":               None,
        "Tipo de Credito":           "TDC",
        "VEN_SALCRE":                None,
        "VEN_CANOTO":                df_origen["Nº Credito"].map(dic_vlookup).fillna(0),
        "Telefono":                  df_origen["Tel Fav"],
        "Tel Casa":                  df_origen["Tel Casa"],
        "Tel Oficina":               df_origen["Tel Oficina"],
        "Tel1":                      df_origen["Tel Cel"],
        "Tel2":                      df_origen["Tel CV"],
        "Tel3":                      df_origen["Tel 1"],
        "Tel4":                      df_origen["Tel 2"],
        "Tel5":                      df_origen["Tel 3"],
        "Tel6":                      df_origen["TelCasa"],
        "Tel7":                      df_origen["TelCelularAdicional"],
        "Tel8":                      df_origen["TelOficina"],
    })


def _clean_phone_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normaliza las columnas de telefono:
    - Reemplaza secuencias de puros ceros por un solo "0".
    - Elimina "nan" de texto.
    - Elimina espacios intermedios (ej. "81 1234 5678" -> "8112345678").

    Args:
        df: DataFrame con las columnas de telefono.

    Returns:
        DataFrame con las columnas de telefono normalizadas (in-place).
    """
    for col in PHONE_COLUMNS:
        if col not in df.columns:
            continue
        df[col] = (
            df[col]
            .astype(str)
            .replace(r"^0+$", "0", regex=True)
            .replace("nan", "")
            .str.replace(" ", "", regex=False)
        )
    return df


def _clean_money_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Convierte las columnas monetarias a numerico, rellenando NaN con 0.

    Args:
        df: DataFrame con las columnas monetarias.

    Returns:
        DataFrame con las columnas monetarias como float (in-place).
    """
    for col in MONEY_COLUMNS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df


def _export_to_csv(df: pd.DataFrame) -> Path:
    """
    Exporta el DataFrame a un archivo CSV con nombre basado en la fecha actual.
    Usa encoding utf-8-sig para compatibilidad con Excel.

    Args:
        df: DataFrame listo para exportar.

    Returns:
        Path absoluto al CSV generado.
    """
    fecha = datetime.datetime.now().strftime("%m %d")
    csv_filename = f"CARGA MC COLLECT B1 {fecha}.csv"
    csv_path = Path(csv_filename).resolve()

    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    log.info("Archivo CSV generado: %s", csv_filename)
    return csv_path


# ── Funcion publica ────────────────────────────────────────────────────────────

def process_excel(
    password: str,
    origin_file: str,
    sheet_name: str,
    output_file: Optional[str] = None,  # noqa: F841 — mantenido por compatibilidad
) -> Path:
    """
    Orquesta el pipeline completo de procesamiento del Excel:
    desencriptacion → lectura → mapeo → limpieza → exportacion CSV.

    Args:
        password:    Contrasena del archivo Excel encriptado.
        origin_file: Ruta al archivo Excel original.
        sheet_name:  Nombre de la pestana principal a leer.
        output_file: (Ignorado, mantenido por compatibilidad con llamadas anteriores.)

    Returns:
        Path absoluto al CSV generado.

    Raises:
        ExcelProcessError: Si el archivo no existe o hay un error en el procesamiento.
    """
    temp_path: Optional[Path] = None
    try:
        # 1. Desencriptar
        temp_path = _decrypt_excel(origin_file, password)

        # 2. Leer hoja principal
        log.info("Leyendo hoja '%s'...", sheet_name)
        df_origen = pd.read_excel(temp_path, sheet_name=sheet_name, dtype=str, skiprows=2)

        # 3. Construir lookup (BUSCARV) desde la misma hoja
        log.debug("Construyendo diccionario BUSCARV interno...")
        df_recave = pd.read_excel(temp_path, sheet_name=sheet_name, dtype=str)
        dic_vlookup = _build_lookup(df_recave)

        # 4. Limpiar nombres de columnas
        df_origen.columns = df_origen.columns.str.strip().str.replace(r"\s+", " ", regex=True)

        # 5. Mapear columnas al formato destino
        log.info("Mapeando columnas al formato de destino...")
        df_nuevo = _build_dataframe(df_origen, dic_vlookup)

        # 6. Limpiar datos
        log.debug("Normalizando columnas de telefono...")
        df_nuevo = _clean_phone_columns(df_nuevo)

        log.debug("Convirtiendo columnas monetarias a numerico...")
        df_nuevo = _clean_money_columns(df_nuevo)

        # 7. Exportar
        log.info("Exportando a CSV...")
        csv_path = _export_to_csv(df_nuevo)

        return csv_path

    finally:
        # Limpiar archivo temporal aunque haya ocurrido un error
        if temp_path and temp_path.exists():
            try:
                temp_path.unlink()
                log.debug("Archivo temporal eliminado: %s", temp_path)
            except OSError:
                log.warning("No se pudo eliminar el archivo temporal: %s", temp_path)
