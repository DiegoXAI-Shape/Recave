import config
from excel_processor import process_excel
from web_automation import run_web_automation

def main():
    print("="*60)
    print("1. PROCESANDO EXCEL...")
    print("="*60)
    process_excel(
        password=config.EXCEL_PASSWORD,
        origin_file=config.EXCEL_ORIGIN,
        sheet_name=config.EXCEL_SHEET,
        output_file=config.EXCEL_CLEAN
    )
    
    print("\n" + "="*60)
    print("2. INICIANDO AUTOMATIZACIÓN WEB...")
    print("="*60)
    run_web_automation()

if __name__ == "__main__":
    main()