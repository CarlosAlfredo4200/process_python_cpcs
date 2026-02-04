import pandas as pd
import json

# Ruta del archivo Excel
archivo_excel = "./dataEstudiantes2.xlsx"   # cambia el nombre si es necesario

# Leer Excel
df = pd.read_excel(archivo_excel, dtype=str)

# Limpiar espacios y valores nulos
df = df.fillna("").applymap(lambda x: x.strip())

# Renombrar columnas si es necesario (seguridad)
df.columns = ["nombre", "documentoIdentidad", "grado"]

# Convertir a lista de diccionarios
data_json = df.to_dict(orient="records")

# Guardar archivo JSON
with open("estudiantes.json", "w", encoding="utf-8") as f:
    json.dump(data_json, f, ensure_ascii=False, indent=4)

print("✅ Archivo estudiantes.json generado correctamente")
