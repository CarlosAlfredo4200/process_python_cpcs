import pandas as pd
import unicodedata


# =========================================================
# FUNCIÓN PARA NORMALIZAR NOMBRES
# =========================================================

def normalizar_nombre(valor):
    """
    Convierte el nombre a mayúsculas, elimina tildes
    y elimina espacios repetidos.
    """

    if pd.isna(valor):
        return ""

    valor = str(valor).strip().upper()

    # Eliminar tildes
    valor = unicodedata.normalize("NFD", valor)

    valor = "".join(
        caracter
        for caracter in valor
        if unicodedata.category(caracter) != "Mn"
    )

    # Eliminar espacios repetidos
    return " ".join(valor.split())


# =========================================================
# FUNCIÓN PARA LEER ARCHIVOS COMPLEMENTARIOS
# =========================================================

def cargar_datos_complementarios(
    ruta_archivo,
    nombre_resultado,
    nombre_persona_archivo
):
    """
    Lee todas las hojas del archivo Excel.

    Columna 0: nombre del estudiante.
    Columna 3: dato principal.
    Columna 5: dato alternativo si la columna 3 está vacía.
    """

    columnas_seleccionadas = [0, 3, 5]

    hojas = pd.read_excel(
        ruta_archivo,
        usecols=columnas_seleccionadas,
        sheet_name=None
    )

    lista_datos = []

    for nombre_hoja, df_hoja in hojas.items():

        if df_hoja.empty:
            continue

        if len(df_hoja.columns) < 3:
            print(
                f"Advertencia: la hoja '{nombre_hoja}' "
                f"de '{ruta_archivo}' no tiene las columnas necesarias."
            )
            continue

        primera_columna = df_hoja.columns[0]
        segunda_columna = df_hoja.columns[1]
        tercera_columna = df_hoja.columns[2]

        df_limpio = df_hoja[
            [
                primera_columna,
                segunda_columna,
                tercera_columna
            ]
        ].copy()

        # Cambiar nombres internos de columnas
        df_limpio.columns = [
            nombre_persona_archivo,
            "valor_columna_3",
            "valor_columna_5"
        ]

        # Limpiar las columnas
        for columna in df_limpio.columns:
            df_limpio[columna] = (
                df_limpio[columna]
                .fillna("")
                .astype(str)
                .str.strip()
            )

        # Eliminar filas que no tengan nombre
        df_limpio = df_limpio[
            df_limpio[nombre_persona_archivo].ne("")
        ].copy()

        # Normalizar nombre para comparar archivos
        df_limpio["nombre_normalizado"] = (
            df_limpio[nombre_persona_archivo]
            .apply(normalizar_nombre)
        )

        # Utilizar la columna 3.
        # Si está vacía, utilizar la columna 5.
        df_limpio[nombre_resultado] = (
            df_limpio["valor_columna_3"]
            .where(
                df_limpio["valor_columna_3"].ne(""),
                df_limpio["valor_columna_5"]
            )
        )

        # Guardar hoja de origen
        df_limpio["hoja_origen"] = nombre_hoja

        lista_datos.append(df_limpio)

    if not lista_datos:
        raise ValueError(
            f"No se encontraron registros válidos en: {ruta_archivo}"
        )

    df_resultado = pd.concat(
        lista_datos,
        ignore_index=True
    )

    # Priorizar registros que sí contienen información
    df_resultado["_dato_vacio"] = (
        df_resultado[nombre_resultado].eq("")
    )

    df_resultado = (
        df_resultado
        .sort_values("_dato_vacio")
        .drop_duplicates(
            subset=["nombre_normalizado"],
            keep="first"
        )
        .drop(columns=["_dato_vacio"])
        .reset_index(drop=True)
    )

    return df_resultado


# =========================================================
# 1. LEER ARCHIVO BASE DE ESTUDIANTES
# =========================================================

ruta_estudiantes = (
    "../Datos_base/estudiantes_global_historico _Q10/"
    "Estudiantes__Informacion_familiar.xlsx"
)

# La posición 47 contiene el grupo.
# La posición 49 contiene el año lectivo.
columnas_base = [
    0,
    1,
    2,
    3,
    4,
    5,
    47,
    49
]

df_nombre_base = pd.read_excel(
    ruta_estudiantes,
    usecols=columnas_base
)

# Como se seleccionaron las columnas en ese orden,
# la columna original 47 queda en la posición 6.
nombre_columna_47 = df_nombre_base.columns[6]

print("Columna utilizada como grupo:", nombre_columna_47)


# =========================================================
# LIMPIAR Y FILTRAR AÑO LECTIVO
# =========================================================

df_nombre_base["Año lectivo"] = (
    df_nombre_base["Año lectivo"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.replace(r"\.0$", "", regex=True)
)

df_nombre_base = df_nombre_base[
    df_nombre_base["Año lectivo"].isin([
        "2026",
        "2026 PRE ESCOLAR"
    ])
].copy()


# =========================================================
# CREAR NOMBRE COMPLETO DEL ESTUDIANTE
# =========================================================

columnas_nombre = [
    "Primer apellido",
    "Segundo apellido",
    "Primer nombre",
    "Segundo nombre"
]

for columna in columnas_nombre:
    df_nombre_base[columna] = (
        df_nombre_base[columna]
        .fillna("")
        .astype(str)
        .str.strip()
    )

df_nombre_base["nombre"] = (
    df_nombre_base["Primer apellido"] + " " +
    df_nombre_base["Segundo apellido"] + " " +
    df_nombre_base["Primer nombre"] + " " +
    df_nombre_base["Segundo nombre"]
)

df_nombre_base["nombre"] = (
    df_nombre_base["nombre"]
    .str.replace(r"\s+", " ", regex=True)
    .str.strip()
)


# =========================================================
# CREAR LISTADO PRINCIPAL DE ESTUDIANTES
# =========================================================

df_lista_nombres = (
    df_nombre_base.loc[
        df_nombre_base["nombre"].ne(""),
        [
            "nombre",
            nombre_columna_47
        ]
    ]
    .drop_duplicates()
    .reset_index(drop=True)
)

df_lista_nombres = df_lista_nombres.rename(
    columns={
        nombre_columna_47: "grupo"
    }
)

df_lista_nombres["grupo"] = (
    df_lista_nombres["grupo"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.replace(r"\.0$", "", regex=True)
)

df_lista_nombres["nombre_normalizado"] = (
    df_lista_nombres["nombre"]
    .apply(normalizar_nombre)
)


# =========================================================
# 2. LEER DATOS DE IGLESIAS
# =========================================================

ruta_iglesias = (
    "../Datos_base/estudiantes_global_historico _Q10/"
    "data_nombre_iglesia.xlsx"
)

df_data_iglesias = cargar_datos_complementarios(
    ruta_archivo=ruta_iglesias,
    nombre_resultado="iglesia",
    nombre_persona_archivo="nombre_archivo_iglesia"
)


# =========================================================
# 3. LEER DATOS DE PASTORES
# =========================================================

ruta_pastores = (
    "../Datos_base/estudiantes_global_historico _Q10/"
    "datos_nombre_pastor.xlsx"
)

df_data_pastores = cargar_datos_complementarios(
    ruta_archivo=ruta_pastores,
    nombre_resultado="pastor",
    nombre_persona_archivo="nombre_archivo_pastor"
)


# =========================================================
# 4. LEER NÚMEROS DE CONTACTO DE IGLESIAS
# =========================================================

ruta_contactos_iglesia = (
    "../Datos_base/estudiantes_global_historico _Q10/"
    "Datos_numero_contacto_iglesia.xlsx"
)

df_data_contactos_iglesia = cargar_datos_complementarios(
    ruta_archivo=ruta_contactos_iglesia,
    nombre_resultado="numero_contacto_iglesia",
    nombre_persona_archivo="nombre_archivo_contacto"
)

# Eliminar ".0" al final si Excel interpretó el teléfono como número
df_data_contactos_iglesia["numero_contacto_iglesia"] = (
    df_data_contactos_iglesia["numero_contacto_iglesia"]
    .str.replace(r"\.0$", "", regex=True)
)


# =========================================================
# 5. LEER CORREOS ELECTRÓNICOS DE IGLESIAS
# =========================================================

ruta_emails_iglesia = (
    "../Datos_base/estudiantes_global_historico _Q10/"
    "datos_email_iglesia.xlsx"
)

df_data_emails_iglesia = cargar_datos_complementarios(
    ruta_archivo=ruta_emails_iglesia,
    nombre_resultado="email_iglesia",
    nombre_persona_archivo="nombre_archivo_email"
)


# =========================================================
# 6. CREAR CONSOLIDADO Y CRUZAR IGLESIAS
# =========================================================

df_consolidado = df_lista_nombres.merge(
    df_data_iglesias[
        [
            "nombre_normalizado",
            "iglesia"
        ]
    ],
    on="nombre_normalizado",
    how="left"
)


# =========================================================
# 7. CRUZAR NOMBRES DE PASTORES
# =========================================================

df_consolidado = df_consolidado.merge(
    df_data_pastores[
        [
            "nombre_normalizado",
            "pastor"
        ]
    ],
    on="nombre_normalizado",
    how="left"
)


# =========================================================
# 8. CRUZAR NÚMEROS DE CONTACTO
# =========================================================

df_consolidado = df_consolidado.merge(
    df_data_contactos_iglesia[
        [
            "nombre_normalizado",
            "numero_contacto_iglesia"
        ]
    ],
    on="nombre_normalizado",
    how="left"
)


# =========================================================
# 9. CRUZAR CORREOS ELECTRÓNICOS
# =========================================================

df_consolidado = df_consolidado.merge(
    df_data_emails_iglesia[
        [
            "nombre_normalizado",
            "email_iglesia"
        ]
    ],
    on="nombre_normalizado",
    how="left"
)


# =========================================================
# 10. REEMPLAZAR VALORES NO ENCONTRADOS
# =========================================================

columnas_resultado = [
    "iglesia",
    "pastor",
    "numero_contacto_iglesia",
    "email_iglesia"
]

for columna in columnas_resultado:
    df_consolidado[columna] = (
        df_consolidado[columna]
        .fillna("")
        .astype(str)
        .str.strip()
        .replace("", "NO ENCONTRADO")
    )


# =========================================================
# 11. SELECCIONAR COLUMNAS FINALES
# =========================================================

df_consolidado = df_consolidado[
    [
        "nombre",
        "grupo",
        "iglesia",
        "pastor",
        "numero_contacto_iglesia",
        "email_iglesia"
    ]
]


# =========================================================
# 12. EXPORTAR ARCHIVO CONSOLIDADO
# =========================================================

ruta_salida = (
    "../Datos_base/estudiantes_global_historico _Q10/"
    "consolidados_unificado.xlsx"
)

df_consolidado.to_excel(
    ruta_salida,
    index=False
)


# =========================================================
# 13. MOSTRAR ESTADÍSTICAS
# =========================================================

total_estudiantes = len(df_consolidado)

iglesias_encontradas = (
    df_consolidado["iglesia"] != "NO ENCONTRADO"
).sum()

iglesias_no_encontradas = (
    df_consolidado["iglesia"] == "NO ENCONTRADO"
).sum()

pastores_encontrados = (
    df_consolidado["pastor"] != "NO ENCONTRADO"
).sum()

pastores_no_encontrados = (
    df_consolidado["pastor"] == "NO ENCONTRADO"
).sum()

contactos_encontrados = (
    df_consolidado["numero_contacto_iglesia"]
    != "NO ENCONTRADO"
).sum()

contactos_no_encontrados = (
    df_consolidado["numero_contacto_iglesia"]
    == "NO ENCONTRADO"
).sum()

emails_encontrados = (
    df_consolidado["email_iglesia"]
    != "NO ENCONTRADO"
).sum()

emails_no_encontrados = (
    df_consolidado["email_iglesia"]
    == "NO ENCONTRADO"
).sum()

print("\n=========================================")
print("PROCESO TERMINADO CORRECTAMENTE")
print("=========================================")

print(f"\nTotal de estudiantes: {total_estudiantes}")

