import pandas as pd

# Leer Excel forzando tipos correctos
df = pd.read_excel(
    "./dataEstudiantesEP.xlsx",
    dtype={
        "codigo": "string",
        "documento": "string"
    }
)

# Limpiar espacios en apellidos
df["apellidos - Copia"] = (
    df["apellidos - Copia"]
    .str.strip()
)

# Crear columna hermanos
df["hermanos"] = df["apellidos - Copia"].map(
    df["apellidos - Copia"].value_counts()
) > 1

# Eliminar columna auxiliar
df = df.drop(columns=["apellidos - Copia"])

# Exportar a JSON
df.to_json(
    "estudiantes_con_hermanos.json",
    orient="records",
    force_ascii=False,
    indent=4
)

print("✅ JSON limpio, sin .0 en codigo ni documento")
