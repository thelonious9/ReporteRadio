import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

def auditoria_excel():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    archivo = filedialog.askopenfilename(title="Selecciona el Excel")
    root.destroy()

    if not archivo: return

    try:
        # 1. Ver qué hojas existen realmente
        xl = pd.ExcelFile(archivo)
        print(f"Hojas detectadas en el archivo: {xl.sheet_names}")

        # 2. Leer la hoja 'Radio' de forma bruta (sin procesar nada aún)
        # Usamos una técnica para leer TODO el archivo, incluso celdas "sueltas"
        df_sucio = pd.read_excel(archivo, sheet_name='Radio', header=None)
        
        print(f"--- Análisis de la hoja 'Radio' ---")
        print(f"Dimensiones crudas detectadas: {df_sucio.shape}")
        
        # 3. Intentar localizar la fila de encabezados automáticamente
        # Buscamos dónde aparece la palabra "Radiodifusora" o "Día"
        df = pd.read_excel(archivo, sheet_name='Radio')
        
        # 4. Forzar la limpieza de celdas vacías pero mantener los datos
        df = df.dropna(subset=[df.columns[0], df.columns[1]], how='all')

        # 5. Exportar y mostrar conteo por 'Día'
        ruta_json = os.path.join(os.path.dirname(__file__), "data_radio.json")
        df.to_json(ruta_json, orient='records', force_ascii=False, indent=4)

        print(f"--- Resultados del JSON ---")
        print(f"Total de filas guardadas: {len(df)}")
        if 'Día' in df.columns:
            print("Conteo por columna 'Día':")
            print(df['Día'].value_counts())
        else:
            print("¡Ojo! No encontré la columna 'Día'. Las columnas son:")
            print(df.columns.tolist())

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    auditoria_excel()