import pandas as pd
import msoffcrypto
import io
import sys
import tempfile
import os
import shutil
import datetime

def process_excel(password, origin_file, sheet_name, output_file):
    if not password:
        print("ERROR: No se encontró la contraseña. Asegúrate de tener tu archivo .env configurado correctamente.")
        sys.exit(1)

    print(f"Desencriptando y leyendo el archivo: {origin_file}...")
    
    temp_fd, temp_path = tempfile.mkstemp(suffix=".xlsx")
    os.close(temp_fd)

    try:
        with open(origin_file, "rb") as file_in:
            office_file = msoffcrypto.OfficeFile(file_in)
            with open(temp_path, "wb") as file_out:
                office_file.load_key(password=password)
                office_file.decrypt(file_out)
    except FileNotFoundError:
        print(f"ERROR: No se encontró el archivo '{origin_file}'.")
        sys.exit(1)
    except Exception as e:
        # Si falla (ej. si el archivo ya está desencriptado o no es OLE), simplemente lo copiamos
        print("El archivo parece ya estar desencriptado o hubo un fallo en msoffcrypto. Usando archivo directo...")
        shutil.copy(origin_file, temp_path)

    # Leer las dos hojas usando Pandas desde el archivo temporal desencriptado
    df_origen = pd.read_excel(temp_path, sheet_name=sheet_name, dtype=str)
    
    print("Mapeando hoja 'Nombre de la Hoja' internamente para el BUSCARV...")
    df_recave = pd.read_excel(temp_path, sheet_name="Nombre de la Hoja", dtype=str)
    
    # BUSCARV INTERNO EFICIENTE
    col_h = df_recave.columns[7]
    col_be = df_recave.columns[56]
    dic_vlookup = dict(zip(df_recave[col_h], df_recave[col_be]))

    # Limpiar columnas
    df_origen.columns = df_origen.columns.str.strip().str.replace(r'\s+', ' ', regex=True)

    print("Archivo leído. Generando nueva hoja 'Prueba'...")

    # Mapeo de columnas
    columnas_diccionario = {
        "Descrip Depto" : df_origen.apply(lambda row: "Consumo B1" if row.get("Bucket") == "B1" and row.get("Descrip Depto") == "Consumo" else row.get("Descrip Depto"), axis=1),
        "Prod" : df_origen["Prod"],
        "N Credito" : df_origen["Nº Credito"],
        "Cta Che" : df_origen["Cta Che"],
        "Cont" : df_origen["Cont"],
        "N Pago Vdo" : df_origen["Nº Pago Vdo"],
        "Plzo" : df_origen["Plzo"],
        "Nom Cliente": df_origen["Nom Cliente"],
        "Fecha Vto" : None,
        "Amo Dias" : df_origen["Amo Dias"],
        "Dias Vdos" : df_origen["Dias Vencidos"],
        "Gestor Nuevo" : None,
        "Saldos Vdos" : df_origen["SaldoPagarUltimoCorte"],
        "Responsa" : df_origen["Responsa"],
        "Descri Credit" : df_origen["Descp Prodcuto"],
        "VEN_TIPCRE" : df_origen.get("VTO"),
        "Capital" : df_origen["CAPITAL"],
        "Interes Nor" : None,
        "Int Mor" : None,
        "Iva Fac" : None,
        "Gasto Cob" : None,
        "Iva Gtos" : None,
        "Gastos Admvos" : None,
        "ADEUDO TOTAL": None,
        "FINIQUITO CON DESCUENTO" : None,
        "Calle" : df_origen["CALLE"],
        "N Calle" : df_origen["NUM CALLE"],
        "Colonia" : df_origen["Colonia"],
        "Codigo Postal" : df_origen["Codigo Postal"],
        "Municipio" : df_origen["Municipio"],
        "Estado" : df_origen["Estado"],
        "Tipo Auto" : None,
        "FechaInicio" : None,
        "Tipo de Credito" : "TDC",
        "VEN_SALCRE" : None,
        "VEN_CANOTO" : df_origen["Nº Credito"].map(dic_vlookup).fillna(0),
        "Telefono" : df_origen["Tel Fav"],
        "Tel Casa" : df_origen["Tel Casa"],
        "Tel Oficina" : df_origen["Tel Oficina"],
        "Tel1" : df_origen["Tel Cel"],
        "Tel2" : df_origen["Tel CV"],
        "Tel3" : df_origen["Tel 1"],
        "Tel4" : df_origen["Tel 2"],
        "Tel5" : df_origen["Tel 3"],
        "Tel6" : df_origen["TelCasa"],
        "Tel7" : df_origen["TelCelularAdicional"],
        "Tel8" : df_origen["TelOficina"],
    }

    df_nuevo = pd.DataFrame(columnas_diccionario)

    # 1. Quitar secuencias completas de ceros (ej: '0000') dejándolos como '0' solamente
    columnas_telefonos = ["Telefono", "Tel Casa", "Tel Oficina", "Tel1", "Tel2", "Tel3", "Tel4", "Tel5", "Tel6", "Tel7", "Tel8"]
    for col in columnas_telefonos:
        if col in df_nuevo.columns:
            # Reemplaza valores que sean puros ceros por un solo "0"
            df_nuevo[col] = df_nuevo[col].astype(str).replace(r'^0+$', '0', regex=True)
            # Limpiar nan de texto por si acaso
            df_nuevo[col] = df_nuevo[col].replace('nan', '')
            # Quitar espacios intermedios (ej: "81 1234 5678" -> "8112345678")
            df_nuevo[col] = df_nuevo[col].str.replace(' ', '', regex=False)

    # 2. Convertimos el dinero y saldos a número
    columnas_dinero = [
        "Capital", "Responsa", "ADEUDO TOTAL", "Saldos Vdos", "VEN_CANOTO"
    ]
    for col in columnas_dinero:
        if col in df_nuevo.columns:
            df_nuevo[col] = pd.to_numeric(df_nuevo[col], errors='coerce').fillna(0)

    print(f"Exportando los datos directamente a CSV...")
    mes_dia = datetime.datetime.now().strftime("%m %d")
    csv_filename = f"CARGA MC COLLECT B1 {mes_dia}.csv"
    csv_path = os.path.abspath(csv_filename)
    
    # Exportar a CSV con encoding utf-8-sig para perfecta compatibilidad con Excel/Web
    df_nuevo.to_csv(csv_path, index=False, encoding='utf-8-sig')
    
    print(f"EXITO: ¡Archivo CSV generado correctamente como '{csv_filename}'!")
    
    # Limpiar temporal si aún existe
    if os.path.exists(temp_path):
        try:
            os.remove(temp_path)
        except:
            pass
            
    return csv_path
