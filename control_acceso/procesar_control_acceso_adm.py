import pandas as pd

input_columns = [1, 3]

data_empleados = pd.read_excel(
    './data/Control_de_acceso_docentes_y_administrativos.xlsx',
    usecols=input_columns
)

data_empleados = data_empleados.rename(columns={
    'Unnamed: 1': 'Nombre empleado',
    'Unnamed: 3': 'horarios'
})

# Eliminar primeras filas
data_empleados = data_empleados.iloc[5:].reset_index(drop=True)

# Eliminar filas sin Nombre empleado
data_empleados = data_empleados.dropna(subset=["Nombre empleado"])

# Limpiar espacios
data_empleados["Nombre empleado"] = (
    data_empleados["Nombre empleado"]
    .astype(str)
    .str.strip()
)

# Eliminar encabezados repetidos
data_empleados = data_empleados[
    ~data_empleados["Nombre empleado"].str.contains("Fecha", na=False)
]

# Identificar filas que son empleados
# Ejemplo: "Álzate Montoya Sandra Elena - 43743409"
data_empleados["es_empleado"] = data_empleados["Nombre empleado"].str.contains(
    r"\s-\s*",
    regex=True,
    na=False
)

# Guardar empleado actual y rellenar hacia abajo
data_empleados["empleado_completo"] = data_empleados["Nombre empleado"].where(
    data_empleados["es_empleado"]
).ffill()

# Separar nombre y documento
data_empleados[["nombre", "documento"]] = (
    data_empleados["empleado_completo"]
    .str.split(" - ", n=1, expand=True)
)

 

# Las filas que NO son empleado son las fechas
data_empleados["fecha"] = data_empleados["Nombre empleado"].where(
    ~data_empleados["es_empleado"]
)


 
# Extraer la primera hora encontrada en la columna horarios
data_empleados["horarios"] = (
    data_empleados["horarios"]
    .astype(str)
    .str.extract(r"(\d{1,2}:\d{2})")[0]
)
 
# Dejar solo filas de fechas
resultado = data_empleados[
    ~data_empleados["es_empleado"]
].copy()

# Seleccionar columnas finales
resultado = resultado[[
    "nombre",
    "documento",
    "fecha",
    "horarios"
]]






# Exportar a Excel
ruta_salida = "./accesos_limpios.xlsx"
resultado.to_excel(ruta_salida, index=False)



# =========================
# LEER ARCHIVO FINAL
# =========================
df = pd.read_excel("./accesos_limpios.xlsx")

# =========================
# LIMPIAR HORARIOS
# =========================
df["horarios"] = (
    df["horarios"]
    .astype(str)
    .str.strip()
)

# Reemplazar textos vacíos por NaN
df["horarios"] = df["horarios"].replace(
    ["", "nan", "None"],
    pd.NA
)

# =========================
# CONVERTIR HORAS
# =========================
df["hora"] = pd.to_datetime(
    df["horarios"],
    format="%H:%M",
    errors="coerce"
)

# Hora límite
hora_limite = pd.to_datetime("06:21", format="%H:%M")

# =========================
# CREAR CONDICIONES
# =========================

# Llegada tarde
df["llegada_tarde"] = df["hora"] > hora_limite

# Llegada a tiempo
df["llegada_tiempo"] = df["hora"] <= hora_limite

# Sin marcar
df["sin_marcar"] = df["hora"].isna()

# =========================
# AGRUPAR POR EMPLEADO
# =========================
reporte = (
    df.groupby(["nombre", "documento"])
    .agg(
        llegadas_tarde=("llegada_tarde", "sum"),
        llegadas_a_tiempo=("llegada_tiempo", "sum"),
        sin_marcar=("sin_marcar", "sum")
    )
    .reset_index()
)

# =========================
# CALCULAR CUMPLIMIENTO
# =========================

reporte["total_registros"] = (
    reporte["llegadas_tarde"]
    + reporte["llegadas_a_tiempo"]
    + reporte["sin_marcar"]
)

reporte["cumplimiento"] = (
    (
        reporte["llegadas_a_tiempo"]
        / reporte["total_registros"]
    ) * 100
).round(2)

# =========================
# TOP TARDANZAS
# =========================
top_tardanzas = reporte.sort_values(
    by="llegadas_tarde",
    ascending=False
)

# =========================
# EXPORTAR
# =========================
with pd.ExcelWriter("./reporte llegadas cpcs abril administrativos.xlsx",
                    engine="openpyxl") as writer:

    # =========================
    # EXPORTAR REPORTE
    # =========================
    reporte.to_excel(
        writer,
        sheet_name="Reporte General",
        startrow=4,   # 👈 deja espacio para títulos
        index=False
    )

    # =========================
    # ACCEDER A LA HOJA
    # =========================
    workbook = writer.book
    worksheet = writer.sheets["Reporte General"]

    # =========================
    # TITULO PRINCIPAL
    # =========================
    worksheet.merge_cells("A1:F1")

    worksheet["A1"] = "REPORTE LLEGADAS EMPLEADOS CPCS"

    # =========================
    # OBSERVACIÓN
    # =========================
    worksheet.merge_cells("A2:F3")

    worksheet["A2"] = (
        "Los registros incluyen algunos valores negativos debido a que el sistema "
        "no discrimina automáticamente días no laborales, periodos de Semana Santa, "
        "festivos y permisos previamente autorizados, lo que puede generar "
        "diferencias temporales en los resultados reportados."
    )

    # =========================
    # ESTILOS
    # =========================
    from openpyxl.styles import Font, Alignment

    # Título
    worksheet["A1"].font = Font(
        bold=True,
        size=16
    )

    worksheet["A1"].alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    # Observación
    worksheet["A2"].alignment = Alignment(
        wrap_text=True,
        vertical="top"
    )

    # Altura filas
    worksheet.row_dimensions[1].height = 28
    worksheet.row_dimensions[2].height = 60

    # Ajustar ancho columnas
    columnas = {
        "A": 40,
        "B": 18,
        "C": 18,
        "D": 22,
        "E": 15,
        "F": 18
    }

    for col, width in columnas.items():
        worksheet.column_dimensions[col].width = width