import pandas as pd 


input_columns_codeudor = [2,18, 21]
datos_codeudor = pd.read_excel('./Document.xlsx', usecols=input_columns_codeudor)
df_codeudor = pd.DataFrame(datos_codeudor)

 
# Cortar filas desde la 16 en adelante
df_codeudor = df_codeudor.iloc[12:].reset_index(drop=True)

# llenar hacia abajo
pd.set_option('future.no_silent_downcasting', True)
df_codeudor = df_codeudor.ffill().infer_objects(copy=False)

# Eliminar filas con NaN en la columna 'Unnamed: 7'
df_codeudor = df_codeudor.dropna(subset=['Unnamed: 2']).reset_index(drop=True)

# limpiar columas y extraer solo el nombre 
df_codeudor['Nombre estudiante'] = df_codeudor['Unnamed: 2'].astype(str).str.split("\n").str[0]
df_codeudor['identificacion'] = df_codeudor['Unnamed: 2'].astype(str).str.extract(r"(\d{8,12})")
df_codeudor['acudiente'] = df_codeudor['Unnamed: 18'].astype(str).str.split("\n").str[0]
df_codeudor['email'] = df_codeudor['Unnamed: 21'].astype(str).str.split("\n").str[1].str.lower()
df_codeudor['telefono'] = df_codeudor['Unnamed: 21'].astype(str).str.split("-").str[0].str.extract(r"(\d{6,12})")
df_codeudor['celular'] = df_codeudor['Unnamed: 21'].astype(str).str.split("-").str[1].str.extract(r"(\d{8,12})")
 

df_codeudor = df_codeudor.drop_duplicates(['Nombre estudiante'])
 
# Exportar a Excel
print("")
ruta_salida = "./pre-familiares.xlsx"
df_codeudor.to_excel(ruta_salida, index=False)
print(f"Archivo Excel creado en: {ruta_salida}")
# Exportar a Excel


 