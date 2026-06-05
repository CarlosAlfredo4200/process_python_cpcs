import pandas as pd

# ===============================
# 1. Cargar archivo JSON
# ===============================

data_json = pd.read_json('../consolidado_informe.json')

# ===============================
# 2. Limpiar datos básicos
# ===============================

data_json['grupo'] = data_json['grupo'].astype(str).str.strip()
data_json['periodo'] = data_json['periodo'].astype(str).str.strip()

# ===============================
# 3. Renombrar columnas
# ===============================

data_json = data_json.rename(columns={
    "ciencias_naturales": "CN",
    "ciencias_sociales": "CS",
    "civica_y_constitucion": "CC",
    "educacion_artistica": "Ed.Art",
    "educacion_cristiana": "Ed.Cris",
    "educacion_etica": "Ed. Etica",
    "educacion_fisica": "Ed. Fis",
    "idioma_extranjero": "Ing",
    "lengua_castellana": "Leng. Cast",
    "matematicas": "Mat",
    "tecnologia": "Tec"
})

# ===============================
# 4. Configuración
# ===============================

periodo_filtrado = "PERIODO 1"

grupos = [
    '1. A', '1. B',
    '2. A', '2. B',
    '3. A', '3. B',
    '4. A', '4. B',
    '5. A', '5. B', '5. C',
    '6. A', '6. B',
    '7. A', '7. B',
    '8. A', '8. B',
    '9. A', '9. B',
    '10. A', '10. B', '10. B1',
    '11. A', '11. B'
]

areas = [
    "CN",
    "CS",
    "CC",
    "Ed.Art",
    "Ed.Cris",
    "Ed. Etica",
    "Ed. Fis",
    "Ing",
    "Leng. Cast",
    "Mat",
    "Tec"
]

# ===============================
# 5. Crear estructura para gráficas
# ===============================

df_grafica = pd.DataFrame()

for grupo in grupos:

    filtro = data_json[
        (data_json['periodo'] == periodo_filtrado) &
        (data_json['grupo'] == grupo)
    ]

    # Si el grupo no existe, lo omite
    if filtro.empty:
        continue

    promedios = []

    for area in areas:
        if area in filtro.columns:
            promedio = round(float(filtro[area].mean()), 1)
        else:
            promedio = None

        promedios.append(promedio)

    df_grafica[grupo] = promedios

# Colocar las áreas como primera columna
df_grafica.insert(0, "grupo", areas)

# ===============================
# 6. Exportar a Excel
# ===============================

nombre_archivo = "graficas_por_grupo_periodo_1.xlsx"

df_grafica.to_excel(nombre_archivo, index=False)

print("Archivo generado correctamente:")
print(nombre_archivo)

print(df_grafica)