import pandas as pd
from datetime import datetime

"""NOTA: el archivo global debe descargarse : familiares y codeudores
        fechas matriculas :todos
        Incluir informacion acudiente
        Estado: todos
        todos los campos
        """

# entrada de datos
print("")
dia_cita = input("Ingrese el dia de la cita (eje. 9): ").strip()
mes_cita = input("Ingrese el mes de la cita (eje. Agosto): ").strip()
hora = input("Ingrese la hora de la cita (eje. 9 am): ").strip()
lugar = input("Ingrese el lugar de la cita : ").strip()


input_columns =[0,2,5,7,22]
archivo = pd.read_excel('./Morosos al 10 de novienbre  2025.xls', usecols=input_columns)
df = pd.DataFrame(archivo)

input_columns_global = [6,57,58,59,61]
data_global_cpcs = pd.read_excel('./Estudiantes__Informacion_familiar.xlsx', usecols=input_columns_global)
df_global = pd.DataFrame(data_global_cpcs)

df_global = df_global.rename(columns={
    "Número de identificación": "identificacion",
})

df_global['identificacion'] = df_global['identificacion'].astype(str).str.extract(r"(\d{8,12})")
 
 

# Cortar filas desde la 16 en adelante
df = df.iloc[16:].reset_index(drop=True)

 
 


# Eliminar filas con NaN en la columna 'Unnamed: 7'
df = df.dropna(subset=['Unnamed: 7']).reset_index(drop=True)

# llenar hacia abajo
pd.set_option('future.no_silent_downcasting', True)
df = df.ffill().infer_objects(copy=False)

 


# limpiar columas y extraer solo el nombre 
df['Nombre estudiante'] = df['Unnamed: 0'].astype(str).str.split("\n").str[0]
df['acudiente'] = df['Unnamed: 2'].astype(str).str.split("\n").str[0]
df['grado'] = df['Unnamed: 5'].astype(str).str.split(" ").str[0]
df['identificacion'] = df['Unnamed: 0'].astype(str).str.extract(r"(\d{8,12})")

# Renombrar columna
df['deuda total'] = df['Unnamed: 22']
df['pensiones pendientes'] = df['Unnamed: 7']

# configurar la fecha actual 
meses = {
    1: "enero",2: "febrero",3: "marzo",4: "abril",5: "mayo",6: "junio",
    7: "julio",8: "agosto",9: "septiembre",10: "octubre",11: "noviembre",12: "diciembre"
}

hoy = datetime.now()
fecha_actual = f"{hoy.day} {meses[hoy.month]} de {hoy.year}"

# columnas adicionales 
df['dia_cita'] = dia_cita
df['mes_cita'] = mes_cita
df['año_cita'] = 2025
df['administrador'] = "Geovanny Callejas Acevedo"
df['fecha actual'] = fecha_actual
df['lugar'] = lugar
df['hora'] = hora 

df = df[['dia_cita', 'mes_cita','año_cita','administrador','fecha actual','lugar','hora','Nombre estudiante','identificacion', 'acudiente','grado','pensiones pendientes','deuda total']]

# Limpiar columna grado 
df = df[df['grado'] !='Grado']


# Copia en paralelo para el filtro de 50
df_fintro_50 = df

# ✅ Filtro: solo registros con deuda > 800.000
df['deuda total'] = pd.to_numeric(df['deuda total'], errors='coerce')
df = df[df['deuda total'] > 800000].reset_index(drop=True)

# ✅ Filtro: solo registros con deuda > 50.000
df['deuda total'] = pd.to_numeric(df['deuda total'], errors='coerce')
df_fintro_50 = df_fintro_50[df_fintro_50['deuda total'] > 50000].reset_index(drop=True)


df_listado_grupo = (
    df[['Nombre estudiante','acudiente','grado']]
    .drop_duplicates(subset=['Nombre estudiante'])   # elimina repetidos en estudiante
    .sort_values(by='grado')                        # ordena por grado
    .reset_index(drop=True)                         # limpia índices
)



# Agrupar
df_final = df.groupby(
    ['dia_cita', 'mes_cita','año_cita','administrador','fecha actual','lugar','hora','Nombre estudiante','identificacion', 'acudiente', 'grado', 'deuda total'], as_index=False
).agg({
    'pensiones pendientes': lambda x: " - ".join(sorted(set(x), key=list(x).index))
})


# Agrupar
df_final_50 = df_fintro_50.groupby(
    ['dia_cita', 'mes_cita','año_cita','administrador','fecha actual','lugar','hora','Nombre estudiante','identificacion', 'acudiente', 'grado', 'deuda total'], as_index=False
).agg({
    'pensiones pendientes': lambda x: " - ".join(sorted(set(x), key=list(x).index))
})


uni_dfs= pd.merge(df_final,df_global, on='identificacion', how="left")

uni_dfs_50= pd.merge(df_final_50,df_global, on='identificacion', how="left")

informe_cartas_y_juntas = uni_dfs[['dia_cita', 'mes_cita','año_cita','administrador','fecha actual','lugar','hora','Nombre estudiante','identificacion', 'acudiente', 'grado', 'deuda total','pensiones pendientes']]

informe_reporte_morosos = uni_dfs[['identificacion','Nombre estudiante','grado', 'acudiente','Parentesco', 'Teléfono acudiente', 'Celular acudiente', 'Correo electrónico acudiente',  'deuda total']]

informe_reporte_morosos_50 = uni_dfs_50[['identificacion','Nombre estudiante','grado', 'acudiente','Parentesco', 'Teléfono acudiente', 'Celular acudiente', 'Correo electrónico acudiente',  'deuda total']]

# Exportar a Excel
print("")
ruta_salida = "./informe_cartas_y_juntas.xlsx"
informe_cartas_y_juntas.to_excel(ruta_salida, index=False)
print(f"Archivo Excel creado en: {ruta_salida}")
# Exportar a Excel

# Exportar a Excel
print("")
ruta_salida = "./informe_reporte_morosos.xlsx"
informe_reporte_morosos.to_excel(ruta_salida, index=False)
print(f"Archivo Excel creado en: {ruta_salida}")
# Exportar a Excel


print("")
ruta_salida = "./morosos_listado.xlsx"
df_listado_grupo.to_excel(ruta_salida, index=False)
print(f"Archivo Excel creado en: {ruta_salida}")

print("")
ruta_salida = "./informe_reporte_morosos_50.xlsx"
informe_reporte_morosos_50.to_excel(ruta_salida, index=False)
print(f"Archivo Excel creado en: {ruta_salida}")

 