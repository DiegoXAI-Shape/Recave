# Recave — Automatización de Carga de Cartera

## Contexto

Este repositorio es un **sandbox** de una automatización real desarrollada en un entorno de trabajo donde el proceso de carga de cartera se realizaba de forma completamente manual: descargar un archivo Excel protegido, procesarlo, y subir el resultado a un portal web interno.

El proyecto existe como espacio de prueba y documentación técnica de las decisiones tomadas durante el desarrollo, incluyendo los obstáculos encontrados y las alternativas evaluadas.

---

## El problema de la detección de bots

La primera aproximación fue automatizar el portal web directamente con **Selenium**, la librería estándar para control de navegadores en Python. El portal la detectaba como tráfico automatizado y bloqueaba la sesión antes de completar la carga.

La segunda aproximación fue **undetected-chromedriver**, una variante de Selenium diseñada para evadir las capas de seguridad más comunes (fingerprinting de WebDriver, análisis de cabeceras, patrones de comportamiento del cursor). Esta variante tampoco fue suficiente: el portal contaba con mecanismos de detección que superaban lo que `undetected-chromedriver` puede ocultar.

La solución adoptada combina Selenium únicamente para la inyección del archivo al input de tipo `file` —operación que no levanta alertas porque no simula interacción humana visible— con **PyAutoGUI** para el resto de la navegación, operando directamente sobre la pantalla como lo haría un usuario real.

---

## Arquitectura del proyecto

```
Recave/
├── main.py                  # Punto de entrada
├── config.py                # Carga y validación de variables de entorno
├── excel_processor.py       # Desencriptado, procesamiento y exportación del Excel
├── web_automation.py        # Sesión web automatizada (Selenium + PyAutoGUI)
├── exceptions.py            # Jerarquía de excepciones del proyecto
├── logger.py                # Configuración centralizada de logging
├── tools/
│   └── inspect_columns.py   # Herramienta de diagnóstico de columnas del Excel
├── images/                  # Capturas de referencia para PyAutoGUI
├── logs/                    # Archivos de log generados en ejecución
└── requirements.txt
```

### Flujo de ejecución

1. `excel_processor` desencripta el archivo Excel con `msoffcrypto`, lee la hoja de trabajo con `pandas`, mapea las columnas al formato que exige el portal y exporta un CSV.
2. `web_automation` abre Chrome con `undetected-chromedriver`, realiza el login y la navegación usando coordenadas de imagen con PyAutoGUI, inyecta el archivo vía Selenium y espera la confirmación del servidor antes de cerrar sesión.

---

## Dependencias principales

| Librería | Propósito |
|---|---|
| `pandas` | Lectura, transformación y exportación del Excel |
| `msoffcrypto-tool` | Desencriptado del archivo Excel protegido con contraseña |
| `undetected-chromedriver` | Control del navegador con evasión de detección básica |
| `selenium` | Inyección de archivo al input `file` del portal |
| `pyautogui` | Interacción con la interfaz gráfica por coordenadas de imagen |
| `python-dotenv` | Gestión de credenciales mediante archivo `.env` |

---

## Configuración

Crear un archivo `.env` en la raíz del proyecto con las siguientes variables:

```
excel_password=contraseña_del_excel
user=usuario_del_portal
pwd=contraseña_del_portal
```

Ajustar en `config.py` el nombre del archivo Excel, la pestaña a leer y la URL del portal si cambian.

---

## Escalabilidad — Alternativas evaluadas

El pipeline actual requiere intervención manual para iniciar la ejecución. A continuación se describen las alternativas técnicas que se evaluaron para automatizar el proceso de extremo a extremo.

### 1. Automatización por correo electrónico

La cartera llega por correo en la mayoría de los flujos reales. Con librerías como `imaplib` o `exchangelib` es posible conectarse al servidor de correo, filtrar el mensaje con los criterios adecuados (remitente, asunto, adjunto), descargar el archivo Excel directamente y pasar al procesamiento sin intervención. Esta opción elimina la dependencia del paso manual de descarga.

### 2. Tarea programada con Cron

Una vez que el pipeline completo funciona sin interacción, la ejecución puede programarse con `cron` en Linux para que corra automáticamente antes de la medianoche, que es el límite habitual para la carga de cartera del día. Una entrada como la siguiente ejecutaría el script a las 23:00:

```
0 23 * * 1-5 /ruta/al/entorno/bin/python /ruta/al/proyecto/main.py
```

Esta opción no tiene coste adicional y es completamente determinista.

### 3. Orquestación con IA (Gemini)

Una alternativa más avanzada sería usar un modelo como Gemini para orquestar el flujo: interpretar el correo entrante, decidir qué archivo procesar, gestionar errores y notificar el resultado. Técnicamente viable, pero introduce una dependencia externa, coste de API y complejidad operativa que no está justificada cuando el problema se resuelve con programación directa.

La opción más pragmática sigue siendo la combinación de cron + lectura de correo: sin dependencias de servicios externos, sin coste incremental y con comportamiento completamente predecible.

---

## Notas

- El archivo Excel original y los CSVs generados están excluidos del repositorio por contener datos sensibles.
- Las imágenes de referencia de PyAutoGUI en `images/` son capturas del portal tomadas en resolución específica. Si cambia la resolución de pantalla o el diseño del portal, deben retomarse.
- El tiempo de espera tras la carga (`PROCESSING_WAIT_SECONDS`) puede ajustarse en `web_automation.py` según la velocidad de respuesta del servidor.
