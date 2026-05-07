import pandas as pd
import msoffcrypto
import io
import sys
import tempfile
import os
import shutil

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
        "VEN_CANOTO" : df_origen["Nº Credito"].map(dic_vlookup).fillna(""),
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

    # 1. Quitar '0' sueltos en columnas de teléfono
    columnas_telefonos = ["Telefono", "Tel Casa", "Tel Oficina", "Tel1", "Tel2", "Tel3", "Tel4", "Tel5", "Tel6", "Tel7", "Tel8"]
    for col in columnas_telefonos:
        if col in df_nuevo.columns:
            df_nuevo[col] = df_nuevo[col].replace("0", "")

    # 2. Convertimos el dinero y saldos a número
    columnas_dinero = [
        "Capital", "Responsa", "ADEUDO TOTAL", "Saldos Vdos"
    ]
    for col in columnas_dinero:
        if col in df_nuevo.columns:
            df_nuevo[col] = pd.to_numeric(df_nuevo[col], errors='coerce')

    print(f"Insertando hoja 'Prueba' en el archivo original: {origin_file}...")
    with pd.ExcelWriter(temp_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_nuevo.to_excel(writer, index=False, sheet_name="Prueba")

    print(f"Re-encriptando el archivo original: {origin_file}...")
    import win32com.client as win32
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        abs_temp_path = os.path.abspath(temp_path)
        abs_origin_path = os.path.abspath(origin_file)
        
        wb = excel.Workbooks.Open(abs_temp_path)
        # FileFormat=51 es xlOpenXMLWorkbook (.xlsx)
        wb.SaveAs(abs_origin_path, FileFormat=51, Password=password)
        wb.Close(SaveChanges=False)
        print(f"EXITO: ¡Hoja 'Prueba' agregada y archivo '{origin_file}' protegido con contraseña exitosamente!")
    except Exception as e:
        print(f"ADVERTENCIA: No se pudo encriptar el archivo. Detalle: {e}")
        shutil.move(temp_path, origin_file)
    finally:
        if excel:
            excel.Quit()
        # Limpiar temporal si aún existe
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except:
                pass
