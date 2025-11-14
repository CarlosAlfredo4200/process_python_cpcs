import pandas as pd 


def procesarOportunidadesFN(data):
    """ 
    El archivo fuente ruta: comercial CRM - Oportunidades
        Procesar el archivo de Oportunidades / Nuevos en el colegio.
    """
    input_columns = [7,8,9,10,11,32,38,42,61]
    data_Oportunidad = pd.read_excel(data, usecols=input_columns)
    df = pd.DataFrame(data_Oportunidad)

    # 🔹 Limpiar espacios en todas las columnas de texto
    df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)

    # 🔹 Filtros
    df_filtro_estado = df[df['Estado negocio'] == 'Aprobado']

    df_filtro_fecha = df_filtro_estado[
        (df_filtro_estado['Fecha de registro'].dt.year == 2025) &
        (df_filtro_estado['Fecha de registro'].dt.month != 1)
    ]

    df_filtro_reingreso = df_filtro_fecha[df_filtro_fecha['2.6 REINGRESO'] == 'No'].reset_index(drop=True)

    df_filtro_reingreso = df_filtro_reingreso.rename(columns={ 
       'Oportunidad':'Nombre estudiante',
       'Correo electrónico':'Correo',
       '2.6 REINGRESO':'Reingreso',
       'Estado negocio':'Estado'
    })

    # 🔹 Función para invertir el orden del nombre completo
    def invertir_nombre(nombre):
        if pd.isna(nombre):
            return nombre
        partes = nombre.split()
        if len(partes) >= 2:
            # Asumimos que los dos últimos son apellidos
            apellidos = partes[-2:]
            nombres = partes[:-2]
            return ' '.join(apellidos + nombres)
        return nombre

    # 🔹 Aplicar la función a la columna
    df_filtro_reingreso['Nombre estudiante'] = df_filtro_reingreso['Nombre estudiante'].apply(invertir_nombre)
    
    print(f"Oportunidades :",df_filtro_reingreso.shape)