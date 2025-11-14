import pandas as pd
from datetime import datetime

"""NOTA: el archivo global debe descargarse : familiares y codeudores
        fechas matriculas :todos
        Incluir informacion acudiente
        Estado: todos
        todos los campos
"""

# ==============================
# 🔹 LECTURA DE ARCHIVOS
# ==============================

input_columns = [0, 10, 30]
archivo_rosetta = pd.read_excel('./MOROSOS ROSETTA 2025.xls', usecols=input_columns)
df_rosetta = pd.DataFrame(archivo_rosetta)

archivo_antologia = pd.read_excel('./MOROSOS ANTOLOGIA 2025.xls', usecols=input_columns)
df_antologia = pd.DataFrame(archivo_antologia)

# Cortar filas desde la 16 en adelante
df_rosetta = df_rosetta.iloc[16:].reset_index(drop=True)
df_antologia = df_antologia.iloc[16:].reset_index(drop=True)

input_columns_global = [6, 57, 58, 59, 61]
data_global_cpcs = pd.read_excel('../Estudiantes__Informacion_familiar.xlsx', usecols=input_columns_global)
df_global = pd.DataFrame(data_global_cpcs)

df_global = df_global.rename(columns={
    "Número de identificación": "identificacion",
})

df_global['identificacion'] = df_global['identificacion'].astype(str).str.extract(r"(\d{8,12})")

# ==============================
# 🔹 LIMPIEZA DE DATOS
# ==============================

df_rosetta = df_rosetta.dropna(subset=['Unnamed: 0']).reset_index(drop=True)
df_antologia = df_antologia.dropna(subset=['Unnamed: 0']).reset_index(drop=True)

pd.set_option('future.no_silent_downcasting', True)
df_rosetta = df_rosetta.ffill().infer_objects(copy=False)
df_antologia = df_antologia.ffill().infer_objects(copy=False)

patron_exclusion = r'(?:total principal|sede\s*-|principal\s*-|www.q10.com|cliente)'
df_rosetta = df_rosetta[~df_rosetta['Unnamed: 0'].str.contains(patron_exclusion, case=False, na=False, regex=True)].reset_index(drop=True)
df_antologia = df_antologia[~df_antologia['Unnamed: 0'].str.contains(patron_exclusion, case=False, na=False, regex=True)].reset_index(drop=True)

# ==============================
# 🔹 LIMPIAR COLUMNAS
# ==============================

df_rosetta['Nombre estudiante'] = df_rosetta['Unnamed: 0'].astype(str).str.split("\n").str[0]
df_rosetta['identificacion'] = df_rosetta['Unnamed: 0'].astype(str).str.extract(r"(\d{8,12})")
df_rosetta['deuda total'] = df_rosetta['Unnamed: 30']

df_antologia['Nombre estudiante'] = df_antologia['Unnamed: 0'].astype(str).str.split("\n").str[0]
df_antologia['identificacion'] = df_antologia['Unnamed: 0'].astype(str).str.extract(r"(\d{8,12})")
df_antologia['deuda total'] = df_antologia['Unnamed: 30']

df_rosetta = df_rosetta[['Nombre estudiante', 'identificacion', 'deuda total']]
df_antologia = df_antologia[['Nombre estudiante', 'identificacion', 'deuda total']]

# ==============================
# 🔹 UNIR CON DATOS GLOBALES
# ==============================

uni_dfs_rosetta = pd.merge(df_rosetta, df_global, on='identificacion', how="left")
uni_dfs_antologia = pd.merge(df_antologia, df_global, on='identificacion', how="left")

# ==============================
# 🔹 FUNCIÓN PARA EXPORTAR CON ESTILO
# ==============================

def exportar_excel_estilizado(df, ruta_salida, nombre_hoja):
    with pd.ExcelWriter(ruta_salida, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=nombre_hoja)

        workbook = writer.book
        worksheet = writer.sheets[nombre_hoja]

        # Formatos personalizados
        formato_encabezado = workbook.add_format({
            'bold': True,
            'font_color': 'white',
            'bg_color': '#4472C4',
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })

        formato_texto = workbook.add_format({'align': 'left', 'border': 1})
        formato_centrado = workbook.add_format({'align': 'center', 'border': 1})
        formato_dinero = workbook.add_format({
            'num_format': '$#,##0',
            'align': 'right',
            'border': 1
        })

        # Aplicar formato al encabezado
        for col_num, valor in enumerate(df.columns.values):
            worksheet.write(0, col_num, valor, formato_encabezado)

        # Ajuste automático del ancho de columnas
        for i, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, max_len)

        # Formatear columna de dinero si existe
        if 'deuda total' in df.columns:
            col_index = df.columns.get_loc('deuda total')
            worksheet.set_column(col_index, col_index, 15, formato_dinero)

        # Congelar fila de encabezado
        worksheet.freeze_panes(1, 0)

    print(f"✅ Archivo Excel creado y estilizado: {ruta_salida}")


# ==============================
# 🔹 EXPORTAR ARCHIVOS
# ==============================

ruta_rosseta = "./informe_morosos_rosseta.xlsx"
ruta_antologia = "./informe_morosos_antologia.xlsx"

exportar_excel_estilizado(uni_dfs_rosetta, ruta_rosseta, "Rosetta")
exportar_excel_estilizado(uni_dfs_antologia, ruta_antologia, "Antologia")

print("\n🎉 Proceso completado correctamente.")
