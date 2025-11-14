import pandas as pd
import json
from pymongo import MongoClient

# Conexión a MongoDB
client = MongoClient("mongodb+srv://alfredomontoyacorreo2:dbcpcs2024@cluster0.wu0k7uf.mongodb.net/")
db = client["test"]
collection = db["evaluacions"]

# Traer todos los documentos (solo respuestas)
data = list(collection.find({}, {"respuestas": 1, "_id": 0}))

# Convertir a DataFrame y aplanar las respuestas
df = pd.DataFrame(data)
df_exploded = df.explode("respuestas").reset_index(drop=True)

# Normalizar cada respuesta en columnas
df_normalized = pd.json_normalize(df_exploded["respuestas"])

# Agrupar por criterio y valor para contar ocurrencias
conteos = df_normalized.groupby(["criterio", "valor"]).size().reset_index(name="total")

# Exportar el consolidado a Excel
ruta_salida = "./consolidado_evaluaciones.xlsx"
conteos.to_excel(ruta_salida, index=False)

print(f"✅ Archivo Excel consolidado creado en: {ruta_salida}")


 
