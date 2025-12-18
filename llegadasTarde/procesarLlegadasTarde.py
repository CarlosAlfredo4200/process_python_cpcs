import pandas as pd
import re
import json
import os
from pymongo import MongoClient
from dotenv import load_dotenv

# ===============================
# CARGAR VARIABLES DE ENTORNO
# ===============================
load_dotenv()

try:
    # ===============================
    # 1. CARGAR TODAS LAS HOJAS DEL EXCEL
    # ===============================
    excel_file = '../llegadasTarde/datosEntrada/Document.xlsx'
    all_sheets = pd.read_excel(excel_file, sheet_name=None)

    resultados = []

    for sheet_name, data in all_sheets.items():
        data = pd.DataFrame(data)

        # ===============================
        # 2. ELIMINAR FILAS INICIALES
        # ===============================
        data = data.iloc[7:].reset_index(drop=True)

        # ===============================
        # 3. ELIMINAR COLUMNAS INNECESARIAS
        # ===============================
        cols_to_drop = [
            'COLEGIO PANAMERICANO COLOMBOSUECO',
            'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4',
            'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 9',
            'Unnamed: 10', 'Unnamed: 8', 'Unnamed: 5'
        ]
        data = data.drop(columns=[c for c in cols_to_drop if c in data.columns])

        # ===============================
        # 4. RENOMBRAR COLUMNAS
        # ===============================
        data = data.rename(columns={
            "Unnamed: 0": "fechas",
        })

        # ===============================
        # 5. EXTRAER NÚMERO DE IDENTIFICACIÓN
        # ===============================
        texto = str(data.loc[0, "fechas"])
        match = re.search(r"(\d+)$", texto)
        num_identificacion = match.group() if match else None

        # ===============================
        # 6. LIMPIAR DATOS DE FECHAS
        # ===============================
        data = data.dropna().reset_index(drop=True)
        data = data.iloc[2:].reset_index(drop=True)

        # ===============================
        # 7. CONVERTIR FECHAS A ISO 8601
        # ===============================
        fechas = pd.to_datetime(
            data["fechas"],
            errors="coerce"
        ).dropna()

        fechas = fechas.dt.strftime(
            "%Y-%m-%dT%H:%M:%S.000Z"
        ).tolist()

        # ===============================
        # 8. AGREGAR RESULTADO
        # ===============================
        if num_identificacion and fechas:
            resultados.append({
                "num_identificacion": num_identificacion,
                "fechas": fechas
            })

    # ===============================
    # 9. EXPORTAR JSON
    # ===============================
    with open("datos_salida.json", "w", encoding="utf-8") as f:
        json.dump(resultados, f, indent=4, ensure_ascii=False)

    print("✅ Archivo JSON generado con éxito: datos_salida.json")

    # ===============================
    # 10. CONEXIÓN A MONGODB (SEGURA)
    # ===============================
    MONGO_URI = os.getenv("MONGO_URI")
    MONGO_DB = os.getenv("MONGO_DB")
    MONGO_COLLECTION = os.getenv("MONGO_COLLECTION_LLEGADAS")

    if not MONGO_URI:
        raise ValueError("❌ No se encontró MONGO_URI en el archivo .env")

    client = MongoClient(MONGO_URI)
    db = client[MONGO_DB]
    collection = db[MONGO_COLLECTION]

    print("✅ Conectado a MongoDB")

    # ===============================
    # 11. LIMPIAR E INSERTAR
    # ===============================
    collection.delete_many({})
    collection.insert_many(resultados)

    total = collection.count_documents({})
    print(f"✅ Insertados en Mongo: {total} registros")

except Exception as e:
    print("❌ Error general:", e)
