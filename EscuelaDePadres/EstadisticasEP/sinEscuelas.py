import pandas as pd

# Leer archivos
data_primer_semestre = pd.read_excel('./sin_asistencia_primer.xlsx')
data_segundo_semestre = pd.read_excel('./asistentes_segundo_semestre.xlsx')

# Normalizar nombres (quitar espacios extra y mayúsculas/minúsculas)
data_primer_semestre["Nombre"] = data_primer_semestre["Nombre"].str.strip().str.upper()
data_segundo_semestre["nombreEstudiante"] = data_segundo_semestre["nombreEstudiante"].str.strip().str.upper()

# 1. Eliminar duplicados en primer semestre
data_primer_semestre = data_primer_semestre.drop_duplicates(subset=["Nombre"])

# 2. Eliminar los que ya están en el segundo semestre
resultado = data_primer_semestre[
    ~data_primer_semestre["Nombre"].isin(data_segundo_semestre["nombreEstudiante"])
]

# Guardar resultado en Excel
output_file = "./faltantes a escuelas.xlsx"
resultado.to_excel(output_file, index=False)

print(f"Archivo creado: {output_file}")
print(resultado.head())
