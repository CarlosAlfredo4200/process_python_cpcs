import pandas as pd 


input_columns_10A = [2,3,4,11,16]
data_10_A = pd.read_excel('./data_X_ grado/Document.xlsx', usecols=input_columns_10A)

df = pd.DataFrame(data_10_A)

# --- Elimina las primeras 5 filas ---
df = df.iloc[5:].reset_index(drop=True)

# llenar hacia abajo
pd.set_option('future.no_silent_downcasting', True)
df = df.ffill().infer_objects(copy=False)

df_flitrada = df[df['Unnamed: 11'] == 'FN']


# # Eliminar filas con NaN en la columna 'Unnamed: 7'
# df = df.dropna(subset=['Unnamed: 4']).reset_index(drop=True)

# Eliminar duplicados de una columna (por ejemplo 'nombre')
df_flitrada = df_flitrada.drop_duplicates(subset=['Unnamed: 4'])


print(df_flitrada)
# # Exportar a Excel
# print("")
# ruta_salida = "./10_A_data.xlsx"
# df.to_excel(ruta_salida, index=False)
# print(f"Archivo Excel creado en: {ruta_salida}")
# # Exportar a Excel

# # print(df.head())