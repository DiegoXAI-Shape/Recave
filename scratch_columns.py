import pandas as pd
import msoffcrypto
import io
import os
import config

password = config.EXCEL_PASSWORD
archivo = config.EXCEL_ORIGIN

try:
    with open(archivo, "rb") as f:
        file = msoffcrypto.OfficeFile(f)
        file.load_key(password=password)
        decrypted = io.BytesIO()
        file.decrypt(decrypted)
        
    df1 = pd.read_excel(decrypted, sheet_name=0, skiprows=1) # Primera hoja
    decrypted.seek(0)
    df2 = pd.read_excel(decrypted, sheet_name="Nombre de la Hoja")
    
    print("\n" + "="*50)
    print("COLUMNAS EN HOJA PRINCIPAL:")
    print("="*50)
    for col in df1.columns:
        print(f"'{col}'")
        
    print("\n" + "="*50)
    print("COLUMNAS EN 'Nombre de la Hoja':")
    print("="*50)
    for col in df2.columns:
        print(f"'{col}'")

except Exception as e:
    print(f"Error: {e}")
