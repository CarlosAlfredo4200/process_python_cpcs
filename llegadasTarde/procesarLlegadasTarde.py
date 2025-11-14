import pandas as pd
import re
import json
from pymongo import MongoClient

# --- Cargar todas las hojas del archivo ---
excel_file = '../llegadasTarde/datosEntrada/Document.xlsx'
all_sheets = pd.read_excel(excel_file, sheet_name=None)  # sheet_name=None lee todas las hojas

resultados = []

for sheet_name, data in all_sheets.items():
    data = pd.DataFrame(data)

    # --- Elimina las primeras 7 filas ---
    data = data.iloc[7:].reset_index(drop=True)

    # --- Eliminar columnas innecesarias ---
    cols_to_drop = [
        'COLEGIO PANAMERICANO COLOMBOSUECO', 
        'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4', 
        'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 9', 
        'Unnamed: 10','Unnamed: 8', "Unnamed: 5"
    ]
    data = data.drop(columns=[c for c in cols_to_drop if c in data.columns])

    # --- Renombrar columnas ---
    data = data.rename(columns={
        "Unnamed: 0": "fechas",
    })

    # --- Extraer número de identificación de la fila 0 ---
    texto = str(data.loc[0, "fechas"])
    match = re.search(r"(\d+)$", texto)
    num_identificacion = match.group() if match else None

    # --- Limpiar datos de fechas ---
    data = data.dropna().reset_index(drop=True)
    data = data.iloc[2:].reset_index(drop=True)

    # --- Convertir fechas a ISO 8601 con milisegundos y Z ---
    fechas = pd.to_datetime(data["fechas"], errors="coerce").dropna()
    fechas = fechas.dt.strftime("%Y-%m-%dT%H:%M:%S.000Z").tolist()

    # --- Guardar en lista de resultados ---
    if num_identificacion and fechas:
        resultados.append({
            "num_identificacion": num_identificacion,
            "fechas": fechas
        })



# --- Exportar lista completa a JSON ---
with open("datos_salida.json", "w", encoding="utf-8") as f:
    json.dump(resultados, f, indent=4, ensure_ascii=False)
    
    
# Conexión a MongoDB (ajusta el puerto y la URI si es necesario)
client = MongoClient("mongodb+srv://alfredomontoyacorreo2:dbcpcs2024@cluster0.wu0k7uf.mongodb.net/")

with open("./datos_salida.json", "r", encoding="utf-8") as file:
    data = json.load(file)

# Seleccionar la base de datos y colección
db = client["test"]
collection = db["llegadastardes"]
 

collection.delete_many({})  # limpiar todo antes
collection.insert_many(data)
total = collection.count_documents({})
print(f'Insertados en Mongo : {total} registros')

print("Archivo JSON generado con éxito: datos_salida.json")
