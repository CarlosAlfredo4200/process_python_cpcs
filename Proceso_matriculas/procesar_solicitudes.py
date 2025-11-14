import pandas as pd 

""" 
    El archivo fuente ruta: Bienestar institucional-solicitudes 
    Procesar el archivo de Solicitudes / antiguos del colegio.
    Fitrar Estado : Aprobados
"""

def procesarSolicitudesFN(data):
 
    inpu_columns_solicitudes = [0,1,2,3,4,6,7]
    data_solicitudes = pd.read_excel(data, usecols=inpu_columns_solicitudes)
    df_solicitudes = pd.DataFrame(data_solicitudes)

    # 🔹 Limpiar espacios en todas las columnas de texto
    df_solicitudes = df_solicitudes.apply(lambda x: x.str.strip() if x.dtype == "object" else x)

    # 🔹 Filtros
    df_filtro_estado_solicitudes = df_solicitudes[df_solicitudes['Estado'] == 'Aprobada']

    df_filtro_fecha_solicitudes = df_filtro_estado_solicitudes[
        (df_filtro_estado_solicitudes['Fecha de la solicitud'].dt.year == 2025) &
        (df_filtro_estado_solicitudes['Fecha de la solicitud'].dt.month != 1)
    ]
    
    # print(df_filtro_fecha_solicitudes.head())
    print(f"Solicitudes :",df_filtro_fecha_solicitudes.shape)
