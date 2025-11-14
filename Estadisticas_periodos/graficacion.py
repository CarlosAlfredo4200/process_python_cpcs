import pandas as pd

# --- 1️⃣ Leer el archivo original ---
df = pd.read_excel("consolidados_data.xlsx")

# --- 2️⃣ Asegurar que la columna "Grupo" existe ---
if 'Grupo' not in df.columns:
    raise ValueError("No se encontró la columna 'Grupo' en el archivo Excel")

# --- 3️⃣ Crear un archivo Excel con varias hojas ---
with pd.ExcelWriter("consolidados_transformado.xlsx", engine='openpyxl') as writer:
    # Iterar sobre cada grupo
    for grupo, datos in df.groupby("Grupo"):
        # Eliminar la columna 'Grupo' para transponer solo las notas
        datos = datos.drop(columns=["Grupo"])

        # Tomar la primera fila (suponiendo que solo hay una fila por grupo)
        fila = datos.iloc[0]

        # Crear DataFrame vertical (mantiene orden correcto)
        df_vertical = pd.DataFrame({
            "Materia": fila.index,
            "Valor": fila.values
        })

        # Escribir en una hoja con el nombre del grupo
        df_vertical.to_excel(writer, sheet_name=str(grupo), index=False)

print("✅ Archivo generado con éxito: consolidados_transformado.xlsx")
