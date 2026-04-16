import pandas as pd


def procesarAcumulado_1_9(ruta_sabana, año):
    input_columns = ['Código de matrícula', 'Tipo de identificación', 'Número de identificación', 'Folio']
    archivo_estudiante = pd.read_excel(
        '../Datos_base/estudiantes_global_historico _Q10/Estudiantes__Informacion_familiar.xlsx',
        usecols=input_columns
    )

    df_data_estudiantes = pd.DataFrame(archivo_estudiante)

    input_columns_sabana = [
        'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4',
        'Unnamed: 9', 'Unnamed: 11', 'Unnamed: 13', 'Unnamed: 15',
        'Unnamed: 17', 'Unnamed: 19', 'Unnamed: 21', 'Unnamed: 23',
        'Unnamed: 24', 'Unnamed: 25', 'Unnamed: 26', 'Unnamed: 28'
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
        'Unnamed: 9': 'cn',
        'Unnamed: 11': 'cs',
        'Unnamed: 13': 'cc',
        'Unnamed: 15': 'art',
        'Unnamed: 17': 'edc',
        'Unnamed: 19': 'ede',
        'Unnamed: 21': 'edf',
        'Unnamed: 23': 'humanidades',
        'Unnamed: 24': 'Ing',
        'Unnamed: 25': 'cast',
        'Unnamed: 26': 'mat',
        'Unnamed: 28': 'tec'
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
    df_resultado['nivel'] = 'pendiente'
    df_resultado['Promovido'] = 'Sí'
    df_resultado['Observación'] = 'Promovido'
    df_resultado['promedio'] = 4.1
    df_resultado['Fisica'] = 0
    df_resultado['Quimica'] = 0
    df_resultado['cp'] = 0
    df_resultado['filosofia'] = 0
    
    df_resultado['estudiante']=df_resultado['estudiante'].str.upper()

    df_resultado = df_resultado[df_resultado['periodo'] == 'Acumulado'].copy()

    df_resultado['grupo'] = df_resultado['grupo'].astype(str).str.strip()

    reemplazos_grupo = {
        'PRIMERO A': '1 A',
        'PRIMERO B': '1 B',
        'PRIMERO C': '1 C',
        'SEGUNDO A': '2 A',
        'SEGUNDO B': '2 B',
        'SEGUNDO C': '2 C',
        'TERCERO A': '3 A',
        'TERCERO B': '3 B',
        'TERCERO C': '3 C',
        'CUARTO A': '4 A',
        'CUARTO B': '4 B',
        'CUARTO C': '4 C',
        'QUINTO A': '5 A',
        'QUINTO B': '5 B',
        'QUINTO C': '5 C',
        'SEXTO A': '6 A',
        'SEXTO B': '6 B',
        'SEXTO C': '6 C',
        'SEPTIMO A': '7 A',
        'SEPTIMO B': '7 B',
        'SEPTIMO C': '7 C',
        'OCTAVO A': '8 A',
        'OCTAVO B': '8 B',
        'OCTAVO C': '8 C',
        'NOVENO A': '9 A',
        'NOVENO B': '9 B',
        'NOVENO C': '9 C',
    }
    df_resultado['grupo'] = df_resultado['grupo'].replace(reemplazos_grupo)

    df_resultado['Tipo de identificación'] = df_resultado['Tipo de identificación'].replace({
         'Tarjeta de Identidad': 'T.I.',
        'Registro Civil de Nacimiento': 'R.C.N.',
        'Cédula de Ciudadanía':'C.C.',
        'Cédula ó Identificación de Extranjería':'C.E.',
        'Permiso por Protección Temporal':'PPT',
    })

    df_resultado.loc[
        df_resultado['grupo'].str.startswith(('1', '2', '3', '4', '5')),
        'nivel'
    ] = 'Básica Primaria'

    df_resultado.loc[
        df_resultado['grupo'].str.startswith(('6', '7', '8', '9')),
        'nivel'
    ] = 'Básica Secundaria'

    ordenar_columnas = [
        'Código de matrícula', 'Folio', 'estudiante', 'Tipo de identificación',
        'Número de identificación', 'año', 'grupo', 'nivel', 'Promovido',
        'Observación', 'promedio', 'cn', 'Fisica', 'Quimica', 'cp', 'cs',
        'cc', 'art', 'edc', 'ede', 'edf', 'filosofia', 'humanidades',
        'Ing', 'cast', 'mat', 'tec'
    ]

    df_resultado = df_resultado[ordenar_columnas]

    return df_resultado