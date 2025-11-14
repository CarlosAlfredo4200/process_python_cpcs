import pandas as pd
from pymongo import MongoClient

# --- Conexión a MongoDB ---
def conectar_mongo():
    client = MongoClient("mongodb+srv://alfredomontoyacorreo2:dbcpcs2024@cluster0.wu0k7uf.mongodb.net/")
    db = client["test"]
    return db["studentnotes"]


# --- Función para calcular promedios por grupo ---
def calcular_promedios_por_grupo(df, grupo_objetivo):
    areas = [
         'lengua_castellana',
        'matematicas',
        'ciencias_naturales',
        'ciencias_sociales', 
        'idioma_extranjero' 
    ]

    periodos = ["PERIODO 1", "PERIODO 2", "PERIODO 3", "PERIODO 4"]
    resultados = []

    for periodo in periodos:
        filtro = df[(df["grupo"] == grupo_objetivo) & (df["periodo"] == periodo)].copy()

        if filtro.empty:
            promedios_dict = {area: 0 for area in areas}
        else:
            promedios_dict = {area: filtro[area].mean() for area in areas}

        promedios_dict["Grupo"] = grupo_objetivo
        promedios_dict["Periodo"] = periodo
        resultados.append(promedios_dict)

    promedios_final = pd.DataFrame(resultados)
    promedios_final = promedios_final[["Grupo", "Periodo"] + areas]
    return promedios_final
