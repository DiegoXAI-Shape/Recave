"""
tools/inspect_columns.py
========================
Herramienta de diagnostico: muestra las columnas disponibles en el
archivo Excel del proyecto para ayudar a mapear o depurar.

Uso:
    python tools/inspect_columns.py
    python tools/inspect_columns.py --file otro_archivo.xlsx --sheet "Otra Hoja"
"""
import argparse
import io
import sys
from pathlib import Path

# Asegurar que el modulo raiz del proyecto sea importable desde tools/
sys.path.insert(0, str(Path(__file__).parent.parent))

import msoffcrypto
import pandas as pd
import config


def inspect(file: str, sheet: str) -> None:
    """
    Desencripta el Excel y muestra las columnas de las dos hojas principales.

    Args:
        file:  Ruta al archivo Excel.
        sheet: Nombre de la hoja secundaria a inspeccionar.
    """
    file_path = Path(file)
    if not file_path.exists():
        print(f"[ERROR] Archivo no encontrado: {file}", file=sys.stderr)
        sys.exit(1)

    print(f"Leyendo: {file_path.name}  |  Hoja: {sheet}\n")

    with file_path.open("rb") as f:
        office_file = msoffcrypto.OfficeFile(f)
        office_file.load_key(password=config.EXCEL_PASSWORD)
        decrypted = io.BytesIO()
        office_file.decrypt(decrypted)

    decrypted.seek(0)
    df_main = pd.read_excel(decrypted, sheet_name=0, skiprows=1)

    decrypted.seek(0)
    df_sheet = pd.read_excel(decrypted, sheet_name=sheet)

    separator = "=" * 50

    print(f"\n{separator}")
    print("COLUMNAS EN HOJA PRINCIPAL (indice 0):")
    print(separator)
    for i, col in enumerate(df_main.columns):
        print(f"  [{i:>3}]  {col!r}")

    print(f"\n{separator}")
    print(f"COLUMNAS EN '{sheet}':")
    print(separator)
    for i, col in enumerate(df_sheet.columns):
        print(f"  [{i:>3}]  {col!r}")


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Muestra las columnas del archivo Excel de Recave.",
    )
    parser.add_argument(
        "--file", "-f",
        default=config.EXCEL_ORIGIN,
        help=f"Ruta al archivo Excel (default: {config.EXCEL_ORIGIN})",
    )
    parser.add_argument(
        "--sheet", "-s",
        default=config.EXCEL_SHEET,
        help=f"Nombre de la hoja a inspeccionar (default: {config.EXCEL_SHEET})",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = _parse_args()
    inspect(file=args.file, sheet=args.sheet)
