import os
import sys
import math
import pandas as pd
from dotenv import load_dotenv
from pymongo import MongoClient
from pymongo.errors import PyMongoError

def clean_str(x):
    """Limpia texto: convierte NaN -> '', recorta espacios y colapsa dobles espacios."""
    if x is None:
        return ""
    if isinstance(x, float) and math.isnan(x):
        return ""
    s = str(x).strip()
    s = " ".join(s.split())
    return s

def main():
    load_dotenv()

    MONGO_URI = os.getenv("MONGO_URI")
    MONGO_DB = os.getenv("MONGO_DB", "test")

    if not MONGO_URI:
        print("ERROR: No se encontró MONGO_URI en el .env")
        sys.exit(1)

    # ====== AJUSTA ESTO ======
    excel_path = "./DataSertificadosEstudia.xlsx"   # <-- Cambia por el nombre real del archivo
    sheet_name = 0                          # 0 = primera hoja, o pon el nombre: "Hoja1"
    collection_name = "certificadosestudio2025"
    # =========================

    # 1) Leer Excel
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str)  # dtype=str evita que cambie tipos
    except Exception as e:
        print(f"ERROR leyendo Excel: {e}")
        sys.exit(1)

    # 2) Normalizar nombres de columnas (por si vienen con tildes/espacios)
    df.columns = [c.strip() for c in df.columns]

    # Columnas esperadas (según lo que mostraste)
    # Nota: "Número_de_identificación" puede venir como "Número_de_identificación" o "Numero_de_identificacion"
    posibles_id = ["Número_de_identificación", "Numero_de_identificacion", "Número_de_identificacion", "Numero_de_identificación"]
    col_id = next((c for c in posibles_id if c in df.columns), None)

    required = ["NOMBRE", "Tipo_de_documento", "Grado", "SECCION"]
    faltantes = [c for c in required if c not in df.columns]
    if col_id is None:
        faltantes.append("Número_de_identificación")

    if faltantes:
        print("ERROR: faltan columnas en el Excel:", faltantes)
        print("Columnas encontradas:", list(df.columns))
        sys.exit(1)

    # 3) Limpieza de campos
    df["NOMBRE"] = df["NOMBRE"].apply(clean_str)
    df["Tipo_de_documento"] = df["Tipo_de_documento"].apply(clean_str)
    df[col_id] = df[col_id].apply(clean_str)         # identificación como TEXTO
    df["Grado"] = df["Grado"].apply(clean_str)
    df["SECCION"] = df["SECCION"].apply(clean_str)

    # 4) Quitar filas vacías (sin nombre o sin identificación)
    df = df[(df["NOMBRE"] != "") & (df[col_id] != "")].copy()

    # 5) Convertir a documentos para Mongo
    docs = []
    for _, row in df.iterrows():
        doc = {
            "NOMBRE": row["NOMBRE"],
            "Tipo_de_documento": row["Tipo_de_documento"],
            "Número_de_identificación": row[col_id],   # queda como string
            "Grado": row["Grado"],
            "SECCION": row["SECCION"],
            
        }
        docs.append(doc)

    if not docs:
        print("No hay registros válidos para insertar (revisa el Excel).")
        sys.exit(0)

    # 6) Conectar a Mongo y cargar
    try:
        client = MongoClient(MONGO_URI)
        db = client[MONGO_DB]
        col = db[collection_name]

        # LIMPIAR colección completa
        deleted = col.delete_many({})
        print(f"✅ Colección '{collection_name}' limpiada. Documentos eliminados: {deleted.deleted_count}")

        # Insertar nuevos
        result = col.insert_many(docs, ordered=False)
        print(f"✅ Insertados: {len(result.inserted_ids)} documentos en '{collection_name}'")

    except PyMongoError as e:
        print(f"ERROR MongoDB: {e}")
        sys.exit(1)
    finally:
        try:
            client.close()
        except Exception:
            pass

if __name__ == "__main__":
    main()