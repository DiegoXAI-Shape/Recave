import pyautogui
import pyautogui
import time
import undetected_chromedriver as uc
import os
from selenium.webdriver.common.by import By
from config import USUARIO_WEB, PASSWORD_WEB, URL_DESTINO

def run_web_automation(archivo_a_subir):
    print("Iniciando el navegador web...")

    opciones = uc.ChromeOptions()
    # Desactivar el popup molesto de "Guardar contraseña"
    opciones.add_experimental_option("prefs", {
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False
    })
    # opciones.add_argument('--headless') # Descomenta para correr oculto

    def esperar_y_hacer_clic(ruta_imagen, timeout=15, confidence=0.8, mover_solo=False):
        """Busca una imagen repetidamente hasta que la encuentra o se acaba el tiempo"""
        print(f"Buscando {ruta_imagen}...")
        inicio = time.time()
        while time.time() - inicio < timeout:
            try:
                coords = pyautogui.locateCenterOnScreen(ruta_imagen, confidence=confidence)
                if coords is not None:
                    # En versiones de PyAutoGUI que devuelven Box o Tuple
                    x, y = coords[0], coords[1]
                    # Pequeña pausa de "reacción humana" al ver el botón
                    time.sleep(0.5)
                    
                    if mover_solo:
                        pyautogui.moveTo(x, y, duration=0.8)
                    else:
                        # Movemos el ratón suavemente hacia el botón
                        pyautogui.moveTo(x, y, duration=0.6)
                        time.sleep(0.2) # mini pausa antes del clic
                        pyautogui.click(x, y)
                    
                    # Pausa obligatoria después de cada acción para darle respiro a la página
                    time.sleep(1.5)
                    return True
            except Exception: # pyscreeze.ImageNotFoundException
                pass
            time.sleep(0.5)
        raise Exception(f"¡Me rendí! No pude encontrar '{ruta_imagen}' en la pantalla después de {timeout} segundos.")

    try:
        # El usuario tiene Chrome 147, así que forzamos esa versión del driver
        driver = uc.Chrome(options=opciones, version_main=147)
        
        print(f"Navegando a: {URL_DESTINO}")
        driver.get(URL_DESTINO)
        
        # ==========================================
        # EXPERIMENTO CON PYAUTOGUI (POR IMÁGENES)
        # ==========================================
        print("Iniciando secuencia de PyAutoGUI por imágenes...")
        
        try:
            # 1. Espera hasta 15 segundos para encontrar el inicio de sesión
            esperar_y_hacer_clic("images/inicio_sesion.png")
            print("¡Clic exitoso!")
            
            # Efecto humano de tecleo
            pyautogui.write(USUARIO_WEB, interval=0.1) 
            pyautogui.press("tab")
            time.sleep(0.5)
            pyautogui.write(PASSWORD_WEB, interval=0.1)
            pyautogui.press("enter")
            
            # 2. Ahora espera pacientemente (hasta 15s) a que la web cargue y aparezca "apartado"
            esperar_y_hacer_clic("images/apartado.png", mover_solo=True)
            print("Ratón posicionado sobre la zona.")

            # 3. Esperar cartera
            esperar_y_hacer_clic("images/cartera.png")
            print("Clic en Cartera exitoso.")

            # ==========================================
            # 1. SELECCIONAR PRODUCTO
            # ==========================================
            esperar_y_hacer_clic("images/producto.png")
            esperar_y_hacer_clic("images/seleccion_producto.png")

            # ==========================================
            # 2. SUBIR ARCHIVO CON SELENIUM
            # ==========================================
            print(f"Iniciando la carga del archivo {os.path.basename(archivo_a_subir)} (Selenium)...")
            ruta_archivo = os.path.abspath(archivo_a_subir)
            time.sleep(2) # Pausa técnica para DOM
            
            input_archivo = driver.find_element(By.ID, "ctl00_CPHMaster_AsyncFileUpload1_ctl02")
            input_archivo.send_keys(ruta_archivo)
            print("¡Archivo inyectado! Subiendo en la web...")
            
            # (Opcional) Pausa para asegurar que la web haya procesado la subida 
            # antes de seleccionar el delimitador
            time.sleep(10)

            # ==========================================
            # 3. SELECCIONAR DELIMITADOR
            # ==========================================
            # Al darle clic al delimitador, se abre el menú desplegable
            esperar_y_hacer_clic("images/delimitador.png")
            # Inmediatamente buscamos la coma en ese menú abierto
            esperar_y_hacer_clic("images/coma.png")
            print("Opciones de delimitador y producto seleccionadas con éxito.")

            # ==========================================
            # 4. POSICIONAR EN CARGAR Y ESPERAR
            # ==========================================
            print("Buscando botón de Cargar...")
            # IMPORTANTE: Solo movemos el ratón, NO damos clic para no causar problemas.
            esperar_y_hacer_clic("images/cargar.png", mover_solo=False)
            print("Ratón en posición sobre Cargar. Esperando 6 minutos por seguridad...")
            
            # Esperamos 7 minutos (420 segundos)
            time.sleep(2 * 60)
            print("Tiempo de espera finalizado.")

            # ==========================================
            # 5. CERRAR SESIÓN
            # ==========================================
            esperar_y_hacer_clic("images/cerrar_sesion.png")
            print("Cerrar sesión exitoso.")

        except Exception as e:
            print("No pude encontrar la imagen en la pantalla. Detalle:", e)
        
        print("\n============================================================")
        print("¡Automatización terminada!")
        input("Presiona la tecla ENTER en esta consola negra cuando quieras cerrar el navegador... ")
        print("Cerrando navegador limpiamente...")
        
        # Al cerrar el navegador explícitamente intentamos evitar el error
        try:
            driver.quit()
        except Exception:
            pass
            
    except Exception as e:
        print(f"Error al abrir el navegador: {e}")
