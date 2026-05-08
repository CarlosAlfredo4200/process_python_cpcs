import pandas as pd


def procesarAcumulado_10_11(ruta_sabana, año):
    input_columns = ['Código de matrícula', 'Tipo de identificación', 'Número de identificación', 'Folio']
    archivo_estudiante = pd.read_excel(
        '../Datos_base/estudiantes_global_historico _Q10/Estudiantes__Informacion_familiar.xlsx',
        usecols=input_columns
    )

    df_data_estudiantes = pd.DataFrame(archivo_estudiante)

 
    input_columns_sabana = [
        'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4',
    'Unnamed: 6', 'Unnamed: 9', 'Unnamed: 10', 'Unnamed: 11',
    'Unnamed: 12', 'Unnamed: 14', 'Unnamed: 16', 'Unnamed: 18',
    'Unnamed: 20', 'Unnamed: 22', 'Unnamed: 24', 'Unnamed: 26',
    'Unnamed: 28', 'Unnamed: 29', 'Unnamed: 30', 'Unnamed: 31',
    'Unnamed: 33'
    ]

    archivo_sabana = pd.read_excel(
        ruta_sabana,
        usecols=input_columns_sabana
    )

    archivo_sabana = archivo_sabana.iloc[8:].reset_index(drop=True)
    df_data_sabanas = pd.DataFrame(archivo_sabana)

    df_data_estudiantes['Código de matrícula'] = pd.to_numeric(
        df_data_estudiantes['Código de matrícula'],
        errors='coerce'
    ).astype('Int64')

    df_data_sabanas = df_data_sabanas.rename(columns={
        'Unnamed: 1': 'grupo',
        'Unnamed: 2': 'codigo',
        'Unnamed: 3': 'estudiante',
        'Unnamed: 4': 'periodo',
        'Unnamed: 6': 'promedio',
        'Unnamed: 9': 'cn',
        "Unnamed: 10": "Fisica",
        "Unnamed: 11": "Quimica",
        "Unnamed: 12": "cp",
        'Unnamed: 14': 'cs',
        'Unnamed: 16': 'cc',
        'Unnamed: 18': 'art',
        'Unnamed: 20': 'edc',
        'Unnamed: 22': 'ede',
        'Unnamed: 24': 'edf',
        'Unnamed: 26': 'filosofia',
        'Unnamed: 28': 'humanidades',
        'Unnamed: 29': 'Ing',
        'Unnamed: 30': 'cast',
        'Unnamed: 31': 'mat',
        'Unnamed: 33': 'tec'
    })

    df_data_sabanas['codigo'] = pd.to_numeric(
        df_data_sabanas['codigo'],
        errors='coerce'
    ).astype('Int64')

    df_data_estudiantes = df_data_estudiantes.dropna(subset=['Código de matrícula'])
    df_data_sabanas = df_data_sabanas.dropna(subset=['codigo'])

    df_resultado = pd.merge(
        df_data_estudiantes,
        df_data_sabanas,
        left_on='Código de matrícula',
        right_on='codigo',
        how='inner'
    )

    df_resultado['año'] = año
    
    df_resultado['Promovido'] = 'Sí'
    df_resultado['Observación'] = 'Promovido'
    df_resultado['nivel'] = 'Media Académica'
    # df_resultado['promedio'] = 4.1
    # df_resultado['Fisica'] = 0
    # df_resultado['Quimica'] = 0
    # df_resultado['cp'] = 0
    # df_resultado['filosofia'] = 0
    
    df_resultado['estudiante']=df_resultado['estudiante'].str.upper()

    df_resultado = df_resultado[df_resultado['periodo'] == 'Acumulado'].copy()

    df_resultado['grupo'] = df_resultado['grupo'].astype(str).str.strip()

    reemplazos_grupo = {
    'DECIMO A': '10 A',
    'DECIMO B': '10 B',
    'DECIMO C': '10 C',
    
    'ONCE A': '11 A',
    'ONCE B': '11 B',
    'ONCE C': '11 C',
}
    
    
    df_resultado['grupo'] = df_resultado['grupo'].replace(reemplazos_grupo)

    df_resultado['Tipo de identificación'] = df_resultado['Tipo de identificación'].replace({
        'Tarjeta de Identidad': 'T.I.',
        'Registro Civil de Nacimiento': 'R.C.N.',
        'Cédula de Ciudadanía':'C.C.',
        'Cédula ó Identificación de Extranjería':'C.E.',
        'Permiso por Protección Temporal':'PPT',
    })

     
   
    ordenar_columnas = [
        'Código de matrícula', 'Folio', 'estudiante', 'Tipo de identificación',
        'Número de identificación', 'año', 'grupo', 'nivel', 'Promovido',
        'Observación', 'promedio', 'cn', 'Fisica', 'Quimica', 'cp', 'cs',
        'cc', 'art', 'edc', 'ede', 'edf', 'filosofia', 'humanidades',
        'Ing', 'cast', 'mat', 'tec'
    ]

    df_resultado = df_resultado[ordenar_columnas]

    return df_resultado