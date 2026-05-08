import pandas as pd

# Leer Excel
df = pd.read_excel(
    "./dataEstudiantesEP.xlsx",
    dtype={
        "codigo": "string",
        "documento": "string"
    }
)

# Limpiar espacios
df["apellidos - Copia"] = (
    df["apellidos - Copia"]
    .astype(str)
    .str.strip()
)

# Contar apellidos
conteo_apellidos = df["apellidos - Copia"].value_counts()

# Detectar hermanos
df["hermanos"] = df["apellidos - Copia"].map(conteo_apellidos) > 1

# Crear grupo de hermanos
df["grupo_hermanos"] = (
    df.groupby("apellidos - Copia")
    .ngroup()
)

# Solo dejar grupo si hay hermanos
df.loc[df["hermanos"] == False, "grupo_hermanos"] = None

# Función para listar hermanos
def obtener_hermanos(apellido):
    personas = df[df["apellidos - Copia"] == apellido]["nombre"].tolist()
    return ", ".join(personas)

# Crear columna con lista de hermanos
df["lista_hermanos"] = df["apellidos - Copia"].apply(obtener_hermanos)

# Eliminar columna auxiliar
df = df.drop(columns=["apellidos - Copia"])

# =========================
# EXPORTAR JSON
# =========================
df.to_json(
    "estudiantes_con_hermanos.json",
    orient="records",
    force_ascii=False,
    indent=4
)

# =========================
# EXPORTAR EXCEL
# =========================
df.to_excel(
    "estudiantes_con_hermanos.xlsx",
    index=False
)

print("✅ JSON y Excel generados correctamente")
