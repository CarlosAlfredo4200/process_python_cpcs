import pandas as pd

def procesarNotas_1_9(archivo):
    """
    Procesa el archivo de notas:
    - Quita filas y columnas innecesarias
    - Renombra columnas
    - Filtra por periodo
    - Ajusta tipos de datos
    - Reemplaza nombres de grados
    - Aplica reglas de observaciones y metas
    Retorna un DataFrame limpio
    """

    # --- Cargar archivo ---
    data_PI = pd.read_excel(archivo)

    # --- Elimina las primeras 8 filas ---
    data_sin_cabecera = data_PI.iloc[8:].reset_index(drop=True)

    # --- Adicionar las columnas faltantes ---
    cols_numericas = ["fisica", "quimica", "ciencias_politicas_economicas", "filosofia"]

    for col in cols_numericas:
        if col in data_sin_cabecera.columns:  # 🔑 si la columna ya existe
            data_sin_cabecera[col] = (
                data_sin_cabecera[col]
                .astype(str)
                .str.replace(",", ".", regex=False)
                .astype(float)
            )
        else:
            # Si la columna no existe, la creamos con valor entero 0
            data_sin_cabecera[col] = 0

    # Convertir explícitamente las nuevas columnas a enteros
    data_sin_cabecera[cols_numericas] = data_sin_cabecera[cols_numericas].astype(int)

    # --- Renombrar columnas ---
    data_sin_cabecera = data_sin_cabecera.rename(columns={
        "Unnamed: 1": "grupo",
        "Unnamed: 2": "codigo",
        "Unnamed: 3": "nombre",
        "Unnamed: 4": "periodo",
         
        "Unnamed: 6": "promedio",
        "Unnamed: 9": "ciencias_naturales",
        "Unnamed: 11": "ciencias_sociales",
        "Unnamed: 13": "civica_y_constitucion",
        "Unnamed: 15": "educacion_artistica",
        "Unnamed: 17": "educacion_cristiana",
        "Unnamed: 19": "educacion_etica",
        "Unnamed: 21": "educacion_fisica",
        "Unnamed: 24": "idioma_extranjero",
        "Unnamed: 25": "lengua_castellana",
        "Unnamed: 26": "matematicas",
        "Unnamed: 28": "tecnologia"
    })

    # --- Eliminar columnas innecesarias ---
    data_sin_cabecera = data_sin_cabecera.drop(columns=[
        "Unnamed: 7", "Unnamed: 10", "Unnamed: 12", "Unnamed: 14",
        "Unnamed: 16", "Unnamed: 18", "Unnamed: 20", "Unnamed: 22",
        "Unnamed: 23", "Unnamed: 27", "Unnamed: 29", "Unnamed: 30","Unnamed: 5"
    ], errors="ignore")

       # Limpiar espacios y convertir a mayúsculas
    data_sin_cabecera["periodo"] = data_sin_cabecera["periodo"].str.strip().str.upper()

    # Ahora filtra
    filtro_periodo = data_sin_cabecera[
        data_sin_cabecera["periodo"].isin(["ACUMULADO"])
    ].copy()

    # --- Cambio de tipo de datos ---
    filtro_periodo[["grupo", "codigo", "nombre", "periodo"]] = (
        filtro_periodo[["grupo", "codigo", "nombre", "periodo"]].astype(str)
    )

    

    cols_numericas = [
        "promedio",
        "ciencias_naturales",
        "fisica",
        "quimica",
        "ciencias_politicas_economicas",
        "ciencias_sociales",
        "civica_y_constitucion",
        "educacion_artistica",
        "educacion_cristiana",
        "educacion_etica",
        "educacion_fisica",
        "filosofia",
        "idioma_extranjero",
        "lengua_castellana",
        "matematicas",
        "tecnologia"
    ]

    for col in cols_numericas:
        filtro_periodo[col] = (
            filtro_periodo[col]
            .astype(str)
            .str.replace(",", ".", regex=False)
            .astype(float)
        )

    # --- Diccionario de reemplazos ---
    reemplazos = {
        "ONCE": "11.",
        "DECIMO": "10.",
        "NOVENO": "9.",
        "OCTAVO": "8.",
        "SEPTIMO": "7.",
        "SEXTO": "6.",
        "QUINTO": "5.",
        "CUARTO": "4.",
        "TERCERO": "3.",
        "SEGUNDO": "2.",
        "PRIMERO": "1."
    }

    for palabra, numero in reemplazos.items():
        filtro_periodo['grupo'] = filtro_periodo['grupo'].str.replace(palabra, numero, regex=False)

    # --- Reorganizar columnas ---
    orden_columnas = [
        "grupo", "codigo", "nombre", "periodo", "promedio",
        "ciencias_naturales", "fisica", "quimica", "ciencias_politicas_economicas",
        "ciencias_sociales", "civica_y_constitucion", "educacion_artistica",
        "educacion_cristiana", "educacion_etica", "educacion_fisica", "filosofia",
        "idioma_extranjero", "lengua_castellana", "matematicas", "tecnologia"
    ]
    filtro_periodo = filtro_periodo[orden_columnas]

   

    return filtro_periodo.reset_index(drop=True)
