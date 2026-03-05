import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import json

def find_header_row(df_raw):
    """Busca la fila que contiene las palabras clave de los encabezados."""
    keywords = ["Radiodifusora", "Fecha", "Día", "Programa"]
    for i, row in df_raw.iterrows():
        # Convertir toda la fila a string y buscar si alguna palabra clave está presente
        row_str = " ".join(row.astype(str).tolist())
        if any(key in row_str for key in keywords):
            return i
    return 0 # Por defecto la primera fila

def auditoria_excel():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    archivo = filedialog.askopenfilename(title="Selecciona el Excel Reporte_Radio.xlsx")
    
    if not archivo: 
        root.destroy()
        return

    try:
        print(f"Abriendo archivo: {archivo}")
        xl = pd.ExcelFile(archivo)
        
        if 'Radio' not in xl.sheet_names:
            messagebox.showerror("Error", "No se encontró la hoja 'Radio' en el archivo.")
            root.destroy()
            return

        # 1. Leer de forma bruta para encontrar el encabezado
        df_raw = pd.read_excel(archivo, sheet_name='Radio', header=None)
        header_idx = find_header_row(df_raw)
        print(f"Encabezado detectado en la fila: {header_idx}")

        # 2. Leer con el encabezado correcto
        df = pd.read_excel(archivo, sheet_name='Radio', header=header_idx)
        
        # 3. Limpieza básica
        # Eliminar filas donde Radiodifusora sea nula
        if 'Radiodifusora' in df.columns:
            df = df.dropna(subset=['Radiodifusora'])
        
        # 4. Normalización de nombres de columnas (quitar espacios, etc)
        df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]

        # 5. Exportar a JSON
        ruta_json = os.path.join(os.path.dirname(__file__), "data_radio.json")
        
        # Convertir a records y guardar con encoding correcto
        records = df.to_dict(orient='records')
        with open(ruta_json, 'w', encoding='utf-8') as f:
            json.dump(records, f, ensure_ascii=False, indent=4, default=str)

        # 6. Resumen para el usuario
        total_filas = len(df)
        estaciones = df['Radiodifusora'].unique().tolist() if 'Radiodifusora' in df.columns else []
        
        print(f"--- Exportación Exitosa ---")
        print(f"Total de filas: {total_filas}")
        print(f"Estaciones detectadas ({len(estaciones)}): {', '.join(map(str, estaciones))}")
        
        resumen = f"Se exportaron {total_filas} registros.\n\nRadiodifusoras encontradas:\n- " + "\n- ".join(map(str, estaciones))
        messagebox.showinfo("Éxito", resumen)

    except Exception as e:
        print(f"Error crítico: {e}")
        messagebox.showerror("Error", f"Ocurrió un error: {e}")
    
    finally:
        root.destroy()

if __name__ == "__main__":
    auditoria_excel()