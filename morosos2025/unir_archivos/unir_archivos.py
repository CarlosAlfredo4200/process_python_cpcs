import pandas as pd

# ==========================
# ARCHIVOS
# ==========================
archivo1 = "../Resultados/informe_reporte_morosos.xlsx"
archivo2 = "../Resultados/Lista whatsapp.xlsx"

# ==========================
# LEER ARCHIVO 1
# ==========================
df1 = pd.read_excel(archivo1, sheet_name="Sheet1")

# dejar solo las columnas necesarias
df1 = df1[
    [
        "identificacion",
        "Nombre estudiante",
        "grado",
        "acudiente",
        "Parentesco",
        "Teléfono acudiente",
        "deuda total",
    ]
]

# limpiar nombres
df1["Nombre estudiante"] = df1["Nombre estudiante"].str.strip().str.upper()

# ==========================
# LEER TODAS LAS HOJAS ARCHIVO 2
# ==========================
xls = pd.ExcelFile(archivo2)

lista = []

for hoja in xls.sheet_names:
    
    df_temp = pd.read_excel(archivo2, sheet_name=hoja, header=None)

    df_temp = df_temp.iloc[:, [0,1,2]]

    df_temp.columns = [
        "Nombre estudiante",
        "Celular",
        "Correo"
    ]

    df_temp["Nombre estudiante"] = df_temp["Nombre estudiante"].str.strip().str.upper()

    lista.append(df_temp)

df2 = pd.concat(lista, ignore_index=True)

# ==========================
# MERGE
# ==========================
df_final = df1.merge(
    df2,
    on="Nombre estudiante",
    how="left"
)

# ==========================
# GUARDAR RESULTADO
# ==========================
df_final.to_excel("estudiantes_completo.xlsx", index=False)

print("Archivo generado correctamente")