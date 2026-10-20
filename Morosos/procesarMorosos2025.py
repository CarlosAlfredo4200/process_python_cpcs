import pandas as pd
from datetime import datetime


"""
NOTA: El archivo global debe descargarse con:
- Familiares y codeudores
- Fechas de matrícula: todos
- Incluir información del acudiente
- Estado: todos
- Todos los campos
"""


# ============================================================
# DATOS DE LA CITA
# ============================================================

print()

dia_cita = input(
    "Ingrese el día de la cita (ej. 9): "
).strip()

mes_cita = input(
    "Ingrese el mes de la cita (ej. Agosto): "
).strip()

hora = input(
    "Ingrese la hora de la cita (ej. 9 a.m.): "
).strip()

lugar = input(
    "Ingrese el lugar de la cita: "
).strip()


# ============================================================
# LEER ARCHIVO DE MOROSOS
# ============================================================

columnas_morosos = [0, 2, 5, 7, 22]

df = pd.read_excel(
    "./MOROSOS AL 12 DE OCTUBRE 2026.xls",
    usecols=columnas_morosos
)


# ============================================================
# LEER ARCHIVO GLOBAL DE ESTUDIANTES
# ============================================================

columnas_global = [6, 57, 58, 59, 61]

df_global = pd.read_excel(
    "./Estudiantes__Informacion_familiar.xlsx",
    usecols=columnas_global
)


# Renombrar la identificación
df_global = df_global.rename(
    columns={
        "Número de identificación": "identificacion"
    }
)


# Extraer únicamente identificaciones de 8 a 12 dígitos
df_global["identificacion"] = (
    df_global["identificacion"]
    .astype(str)
    .str.extract(r"(\d{8,12})", expand=False)
)


# Eliminar registros sin identificación en el archivo global
df_global = (
    df_global
    .dropna(subset=["identificacion"])
    .drop_duplicates(subset=["identificacion"])
    .reset_index(drop=True)
)


# ============================================================
# LIMPIAR ARCHIVO DE MOROSOS
# ============================================================

# Eliminar las primeras 16 filas
df = df.iloc[16:].reset_index(drop=True)


# Limpiar los nombres de las columnas por seguridad
df.columns = df.columns.astype(str).str.strip()


# Verificar que las columnas necesarias existan
columnas_necesarias = [
    "Unnamed: 0",
    "Unnamed: 2",
    "Unnamed: 5",
    "Unnamed: 7",
    "Unnamed: 22"
]

columnas_faltantes = [
    columna
    for columna in columnas_necesarias
    if columna not in df.columns
]

if columnas_faltantes:
    raise KeyError(
        "No se encontraron estas columnas en el archivo: "
        + ", ".join(columnas_faltantes)
    )


# ============================================================
# CONSERVAR LA COLUMNA 7 ORIGINAL
# ============================================================

# La columna 7 contiene los meses.
# No se debe aplicar ffill() sobre esta columna porque los NaN
# se convertirían incorrectamente en meses repetidos.
pensiones_originales = df["Unnamed: 7"].copy()


# ============================================================
# LLENAR DATOS DEL ESTUDIANTE HACIA ABAJO
# ============================================================

# Estas columnas sí se rellenan porque las filas de meses
# adicionales tienen vacíos los datos del estudiante.
columnas_rellenar = [
    "Unnamed: 0",
    "Unnamed: 2",
    "Unnamed: 5",
    "Unnamed: 22"
]

pd.set_option("future.no_silent_downcasting", True)

df[columnas_rellenar] = (
    df[columnas_rellenar]
    .ffill()
    .infer_objects(copy=False)
)


# Restaurar la columna 7 con sus valores originales
df["Unnamed: 7"] = pensiones_originales


# Eliminar filas que realmente no tienen información en la columna 7
df = df.dropna(subset=["Unnamed: 7"]).reset_index(drop=True)


# ============================================================
# LIMPIAR LA COLUMNA DE PENSIONES
# ============================================================

df["pensiones pendientes"] = (
    df["Unnamed: 7"]
    .astype(str)
    .str.strip()
    .str.title()
)


# Eliminar valores vacíos, puntos o textos que representan nulos
valores_invalidos = [
    "",
    ".",
    "Nan",
    "None",
    "N/A"
]

df = df[
    ~df["pensiones pendientes"].isin(valores_invalidos)
].reset_index(drop=True)


# ============================================================
# EXTRAER DATOS DEL ESTUDIANTE
# ============================================================

df["Nombre estudiante"] = (
    df["Unnamed: 0"]
    .astype(str)
    .str.split("\n")
    .str[0]
    .str.strip()
)


df["acudiente"] = (
    df["Unnamed: 2"]
    .astype(str)
    .str.split("\n")
    .str[0]
    .str.strip()
)


df["grado"] = (
    df["Unnamed: 5"]
    .astype(str)
    .str.split()
    .str[0]
    .str.strip()
)


df["identificacion"] = (
    df["Unnamed: 0"]
    .astype(str)
    .str.extract(r"(\d{8,12})", expand=False)
)


df["deuda total"] = df["Unnamed: 22"]


# Eliminar filas sin identificación
df = df.dropna(subset=["identificacion"]).reset_index(drop=True)


# Eliminar encabezados repetidos dentro del Excel
df = df[
    df["grado"].str.lower() != "grado"
].reset_index(drop=True)


# ============================================================
# FILTRAR ESTUDIANTES CON DOS O MÁS REGISTROS
# ============================================================

# Cuenta cuántas filas con pensión tiene cada estudiante
df["cantidad_pensiones"] = (
    df.groupby("identificacion")["pensiones pendientes"]
    .transform("count")
)


# Conservar únicamente estudiantes con dos o más registros
df = df[
    df["cantidad_pensiones"] >= 1
].reset_index(drop=True)


# Mostrar resultado del filtro en la terminal
print()
print("Registros encontrados después del filtro:")

print(
    df[
        [
            "Nombre estudiante",
            "identificacion",
            "pensiones pendientes",
            "deuda total",
            "cantidad_pensiones"
        ]
    ].to_string(index=False)
)

print()
print(
    "Estudiantes con dos o más registros:",
    df["identificacion"].nunique()
)


# ============================================================
# CONFIGURAR FECHA ACTUAL
# ============================================================

meses = {
    1: "enero",
    2: "febrero",
    3: "marzo",
    4: "abril",
    5: "mayo",
    6: "junio",
    7: "julio",
    8: "agosto",
    9: "septiembre",
    10: "octubre",
    11: "noviembre",
    12: "diciembre"
}

hoy = datetime.now()

fecha_actual = (
    f"{hoy.day} de {meses[hoy.month]} de {hoy.year}"
)


# ============================================================
# AGREGAR DATOS DE LA CITA
# ============================================================

df["dia_cita"] = dia_cita
df["mes_cita"] = mes_cita
df["año_cita"] = 2026
df["administrador"] = "Liliana María García Marín"
df["fecha actual"] = fecha_actual
df["lugar"] = lugar
df["hora"] = hora


# ============================================================
# CREAR LISTADO POR GRUPO
# ============================================================

df_listado_grupo = (
    df[
        [
            "Nombre estudiante",
            "acudiente",
            "grado"
        ]
    ]
    .drop_duplicates(subset=["Nombre estudiante"])
    .sort_values(
        by=["grado", "Nombre estudiante"]
    )
    .reset_index(drop=True)
)


# ============================================================
# AGRUPAR LOS MESES POR ESTUDIANTE
# ============================================================

columnas_agrupacion = [
    "dia_cita",
    "mes_cita",
    "año_cita",
    "administrador",
    "fecha actual",
    "lugar",
    "hora",
    "Nombre estudiante",
    "identificacion",
    "acudiente",
    "grado",
    "deuda total"
]


def unir_pensiones(valores):
    """
    Une los meses sin repetirlos y conserva
    el orden en que aparecen en el archivo.
    """

    pensiones = []

    for valor in valores:
        valor = str(valor).strip()

        if valor and valor not in pensiones:
            pensiones.append(valor)

    return " - ".join(pensiones)


df_final = (
    df.groupby(
        columnas_agrupacion,
        as_index=False,
        dropna=False
    )
    .agg({
        "pensiones pendientes": unir_pensiones
    })
)


# ============================================================
# UNIR CON EL ARCHIVO GLOBAL
# ============================================================

uni_dfs = pd.merge(
    df_final,
    df_global,
    on="identificacion",
    how="left"
)


# ============================================================
# CREAR INFORME DE CARTAS Y JUNTAS
# ============================================================

informe_cartas_y_juntas = uni_dfs[
    [
        "dia_cita",
        "mes_cita",
        "año_cita",
        "administrador",
        "fecha actual",
        "lugar",
        "hora",
        "Nombre estudiante",
        "identificacion",
        "acudiente",
        "grado",
        "deuda total",
        "pensiones pendientes"
    ]
]


# ============================================================
# CREAR INFORME GENERAL DE MOROSOS
# ============================================================

informe_reporte_morosos = uni_dfs[
    [
        "identificacion",
        "Nombre estudiante",
        "grado",
        "acudiente",
        "Parentesco",
        "Teléfono acudiente",
        "Celular acudiente",
        "Correo electrónico acudiente",
        "deuda total",
        "pensiones pendientes"
    ]
]


# ============================================================
# EXPORTAR INFORME DE CARTAS Y JUNTAS
# ============================================================

print()

ruta_cartas = "./informe_cartas_y_juntas_12_oct.xlsx"

informe_cartas_y_juntas.to_excel(
    ruta_cartas,
    index=False
)

print(f"Archivo Excel creado en: {ruta_cartas}")


# ============================================================
# EXPORTAR INFORME GENERAL DE MOROSOS
# ============================================================

print()

ruta_morosos = "./MOROSOS_AL_12_oct_2026_filtrados.xlsx"

informe_reporte_morosos.to_excel(
    ruta_morosos,
    index=False
)

print(f"Archivo Excel creado en: {ruta_morosos}")


# ============================================================
# EXPORTAR LISTADO POR GRUPO
# ============================================================

print()

ruta_listado = "./morosos_listado_12_oct_filtrados.xlsx"

df_listado_grupo.to_excel(
    ruta_listado,
    index=False
)

print(f"Archivo Excel creado en: {ruta_listado}")


# ============================================================
# RESUMEN
# ============================================================

print()
print(f"Total de estudiantes exportados: {len(df_final)}")