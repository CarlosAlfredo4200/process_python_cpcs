import json
import pandas as pd

# ==========================
# Cargar archivo JSON
# ==========================

with open("inventario.json", "r", encoding="utf-8") as f:
    data = json.load(f)

# ==========================
# Convertir JSON a DataFrame
# ==========================

df = pd.json_normalize(data)

print("Columnas encontradas:")
print(df.columns.tolist())

# ==========================
# Filtrar producto específico
# ==========================

producto_buscado = "Tablet Samsung A7"

df_filtrado = df[
    df["id_producto.name"] == producto_buscado
]

# ==========================
# Mostrar resultado
# ==========================

columnas_mostrar = [
    "id_producto.name",
    "id_producto.brand",
    "id_producto.sku",
    "location.nombre",
    "recibido_por",
    "estado",
    "fecha_entrega",
    "observaciones"
]

# Mostrar solo columnas existentes
columnas_existentes = [
    col for col in columnas_mostrar
    if col in df_filtrado.columns
]

print("\nRegistros encontrados:")
print(df_filtrado[columnas_existentes])

# ==========================
# Guardar en Excel
# ==========================

df_filtrado.to_excel(
    "Tablet_Samsung_A7.xlsx",
    index=False
)

print(
    f"\nSe encontraron {len(df_filtrado)} registros."
)
print("Archivo generado: Tablet_Samsung_A7.xlsx")