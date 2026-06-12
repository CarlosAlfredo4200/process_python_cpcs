import pandas as pd  

def procesar_data_media(data):
    data_media = pd.read_excel(data)
    
    data_media = data_media.iloc[8:].reset_index(drop=True)
    
    data_media = data_media.drop(columns=[
        'Unnamed: 0',
        'Unnamed: 7',
        'Unnamed: 8', 
        'Unnamed: 13',
        'Unnamed: 15',
        'Unnamed: 17',
        'Unnamed: 19',
        'Unnamed: 21',
        'Unnamed: 23',
        'Unnamed: 25',
        'Unnamed: 27',
        'Unnamed: 28',
        'Unnamed: 32',
        'Unnamed: 34',
        'Unnamed: 35'
    ], errors="ignore")
    
    data_media = data_media.rename(columns={
        "Unnamed: 1": "grupo",
        "Unnamed: 2": "codigo",
        "Unnamed: 3": "nombre",
        "Unnamed: 4": "periodo",
        "Unnamed: 5": "puesto",
        "Unnamed: 6": "promedio",
        "Unnamed: 9": "ciencias_naturales",
        "Unnamed: 10": "fisica",
        "Unnamed: 11": "quimica",
        "Unnamed: 12": "ciencias_politicas_economicas",
        "Unnamed: 14": "ciencias_sociales",
        "Unnamed: 16": "civica_y_constitucion",
        "Unnamed: 18": "educacion_artistica",
        "Unnamed: 20": "educacion_cristiana",
        "Unnamed: 22": "educacion_etica",
        "Unnamed: 24": "educacion_fisica",
        "Unnamed: 26": "filosofia",
        "Unnamed: 29": "idioma_extranjero",
        "Unnamed: 30": "lengua_castellana",
        "Unnamed: 31": "matematicas",
        "Unnamed: 33": "tecnologia"
    })
    
    data_media["periodo"] = data_media["periodo"].astype(str).str.strip().str.upper()
    
    data_media = data_media[
        data_media["periodo"].isin(["1 PERIODO", "SEGUNDO PERIODO"])
    ]

    data_media[["grupo", "codigo", "nombre", "periodo"]] = (
        data_media[["grupo", "codigo", "nombre", "periodo"]].astype(str)
    )
    
    data_media["grupo"] = (
        data_media["grupo"]
        .str.strip()
        .str.upper()
        .str.replace(r"\s+", " ", regex=True)
    )

    data_media["puesto"] = pd.to_numeric(
        data_media["puesto"], errors="coerce"
    ).astype("Int64")

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
        data_media[col] = (
            data_media[col]
            .astype(str)
            .str.replace(",", ".", regex=False)
        )

        data_media[col] = pd.to_numeric(data_media[col], errors="coerce")

    reemplazos = {
        "ONCE": "11.",
        "DECIMO": "10.",
        "DÉCIMO": "10.",
        "NOVENO": "9.",
        "OCTAVO": "8.",
        "SEPTIMO": "7.",
        "SÉPTIMO": "7.",
        "SEXTO": "6.",
        "QUINTO": "5.",
        "CUARTO": "4.",
        "TERCERO": "3.",
        "SEGUNDO": "2.",
        "PRIMERO": "1."
    }

    for palabra, numero in reemplazos.items(): 
        data_media["grupo"] = data_media["grupo"].str.replace(
            palabra, numero, regex=False
        )

    data_media = data_media.rename(columns={
        "ciencias_naturales": "CN",
        "ciencias_sociales": "CS",
        "civica_y_constitucion": "CC",
        "educacion_artistica": "ED-Art",
        "educacion_cristiana": "ED-Cris",
        "educacion_etica": "ED-Etica",
        "educacion_fisica": "ED-Fis",
        "idioma_extranjero": "ING",
        "lengua_castellana": "LENG-Cast",
        "matematicas": "MAT",
        "tecnologia": "TEC",
        "ciencias_politicas_economicas": "CPE"
    })

    areas = [
        "CN",
        "CS",
        "CC",
        "ED-Art",
        "ED-Cris",
        "ED-Etica",
        "ED-Fis",
        "ING",
        "LENG-Cast",
        "MAT",
        "TEC",
        "CPE"
    ]

    promedios_areas = (
        data_media
        .groupby(["grupo", "periodo"])[areas]
        .mean()
        .round(2)
        .reset_index()
    )

    return promedios_areas