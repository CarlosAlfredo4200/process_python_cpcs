import pandas as pd

"""
Procesa el archivo de cartera de pensiones para extraer información de pagos de matrículas.
Archivo fuente: informes-gestion financiera / cartera_pensiones
"""

def procesarCarteraFN(data):
    # Columnas de interés
    input_columns_cartera = [0, 2, 5, 7, 15, 18]
    df = pd.read_excel(data, usecols=input_columns_cartera)

    # Saltar filas iniciales innecesarias
    df = df.iloc[16:].reset_index(drop=True)

    # Eliminar filas vacías
    df = df.dropna(subset=['Unnamed: 0']).reset_index(drop=True)

    # Llenar hacia abajo valores vacíos
    pd.set_option('future.no_silent_downcasting', True)
    df = df.ffill().infer_objects(copy=False)

    # Renombrar columnas relevantes
    df = df.rename(columns={
        'Unnamed: 7': 'Concepto',
        'Unnamed: 15': 'A pagar',
        'Unnamed: 18': 'Saldo',
    })

    # Extraer datos específicos
    df['Estudiante'] = df['Unnamed: 0'].astype(str).str.split("\n").str[0]
    df['Identificación'] = df['Unnamed: 0'].astype(str).str.extract(r"(\d{8,12})")

    df['Codeudor'] = df['Unnamed: 2'].astype(str).str.split("\n").str[0]
    df['Identificación_codeudor'] = df['Unnamed: 2'].astype(str).str.extract(r"(\d{8,12})")

    df['Año'] = df['Unnamed: 5'].astype(str).str.extract(r"(\d{4})")
    df['Año'] = pd.to_numeric(df['Año'], errors='coerce')

    df['Grado'] = df['Unnamed: 5'].astype(str).str.split(" ").str[0]

    # 🧹 Eliminar filas no deseadas
    eliminar_textos = [
        'Total Principal',
        'Sede - Jornada - Programa',
        'Principal - 7:00 am - 1:00 pm',
        'Estudiante',
        'www.q10.com'
    ]

    # Elimina cualquier fila donde la columna 'Estudiante' contenga uno de esos textos
    patron = '|'.join(eliminar_textos)
    df = df[~df['Estudiante'].str.contains(patron, case=False, na=False)]

    # Filtrar solo registros del año 2025
    df_filtro_Cartera = df[df['Año'] == 2025]

    # Seleccionar columnas finales
    df_filtro_Cartera = df_filtro_Cartera[
        ['Estudiante', 'Identificación', 'Codeudor', 'Identificación_codeudor', 'Año', 'Grado']
    ].reset_index(drop=True)

    print(df_filtro_Cartera)
    return df_filtro_Cartera
