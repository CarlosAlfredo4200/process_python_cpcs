import pandas as pd

columnas_de_entrada = [1, 2, 3, 4, 6, 47]
archivo_base = pd.read_excel(
    '../Datos_base/Esc_Padres/Estudiantes__Informacion_familiar.xlsx',
    usecols=columnas_de_entrada,
    dtype={'Número de identificación': 'string'}
)



cols = ['Primer apellido', 'Segundo apellido', 'Primer nombre', 'Segundo nombre']

# Construir nombreCompleto ignorando partes vacías/NA
archivo_base['nombreCompleto'] = archivo_base[cols] \
    .fillna('') \
    .agg(' '.join, axis=1) \
    .str.strip()

# Opcional: dejar NA donde no hay ninguna parte del nombre
archivo_base['nombreCompleto'] = archivo_base['nombreCompleto'].replace('', pd.NA)

# Eliminar columnas originales si ya no las quieres
archivo_base = archivo_base.drop(columns=cols)

# Si quieres descartar sólo filas sin identificación o sin grupo:
archivo_base = archivo_base.dropna(subset=['Número de identificación', 'Grupo']).reset_index(drop=True)

baseGlobalColegio = archivo_base


# ====================================================


