import pandas as pd
import json
import win32com.client
import os

def refresh_and_export():
    file_path = os.path.abspath("Reporte_Radio.xlsx")
    json_path = "data_radio.json"
    
    # 1. Actualizar Power Query
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(file_path)
        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()
        wb.Save()
        wb.Close()
        excel.Quit()
        print("Excel actualizado.")
    except Exception as e:
        print(f"Error al actualizar Excel: {e}")

    # 2. Procesar Datos con Pandas
    df = pd.read_excel(file_path, sheet_name='Radio')
    
    # --- SOLUCIÓN AL ERROR DE TIMESTAMP ---
    # Convertimos todas las columnas de fecha a formato texto (DD/MM/YYYY)
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime('%d/%m/%Y')
    # ---------------------------------------

    # Normalización de texto
    def clean_text(t):
        if not isinstance(t, str) or pd.isna(t): return "N/A"
        t = t.strip().capitalize()
        mapping = {'Ssph': 'SSPH', 'Difh': 'DIFH', 'Seph': 'SEPH', 'Semot': 'SEMOT', 'Ssh': 'SSH'}
        return mapping.get(t, t)

    cols_to_fix = ['Radiodifusora', 'Programa', 'Dependencia', 'Estatus', 'Clasificación', 'Autor', 'Género']
    print(f"Columnas encontradas: {df.columns.tolist()}")
    for col in cols_to_fix:
        if col not in df.columns:
            print(f"Advertencia: Columna {col} no encontrada. Creando con 'N/A'.")
            df[col] = "N/A"
        df[col] = df[col].apply(clean_text)

    # Asegurar que Recuento sea numérico y llenar vacíos
    df['Recuento'] = pd.to_numeric(df['Recuento'], errors='coerce').fillna(0)
    df = df.fillna("N/A")

    # Exportar a JSON
    data = df.to_dict(orient='records')
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    print(f"JSON generado exitosamente con {len(data)} registros.")

if __name__ == "__main__":
    refresh_and_export()