import pandas as pd
import re
import unicodedata
from procesarAcumuladosQ10 import procesarAcumuladosQ10


def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    valor = str(valor).strip().upper()
    valor = unicodedata.normalize("NFKD", valor).encode("ascii", "ignore").decode("utf-8")
    valor = re.sub(r"\s+", " ", valor)
    return valor


def convertir_numerica(serie):
    return pd.to_numeric(
        serie.astype(str).str.replace(",", ".", regex=False).str.strip(),
        errors="coerce"
    ).fillna(0)


def procesarNotas(
    archivo_notas,
    archivo_acumulados="../Datos_base/estudiantes_global_historico _Q10/Estudiantes__Informacion_familiar.xlsx",
    ruta_salida=None
):
    # =========================
    # 1. PROCESAR ARCHIVO NOTAS
    # =========================
    data_PI = pd.read_excel(archivo_notas)

    data_sin_cabecera = data_PI.iloc[8:].reset_index(drop=True)

    data_sin_cabecera = data_sin_cabecera.drop(columns=[
        'Unnamed: 0',
        'Unnamed: 5',
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

    data_sin_cabecera = data_sin_cabecera.rename(columns={
        "Unnamed: 1": "grupo",
        "Unnamed: 2": "codigo",
        "Unnamed: 3": "nombre",
        "Unnamed: 4": "periodo",
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

    data_sin_cabecera["periodo"] = (
        data_sin_cabecera["periodo"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    filtro_periodo = data_sin_cabecera[
        data_sin_cabecera["periodo"].isin(["ACUMULADO"])
    ].copy()

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
        if col not in filtro_periodo.columns:
            filtro_periodo[col] = 0
        filtro_periodo[col] = convertir_numerica(filtro_periodo[col])

    reemplazos = {
        "ONCE": "11",
        "DECIMO": "10",
        "NOVENO": "9",
        "OCTAVO": "8",
        "SEPTIMO": "7",
        "SEXTO": "6",
        "QUINTO": "5",
        "CUARTO": "4",
        "TERCERO": "3",
        "SEGUNDO": "2",
        "PRIMERO": "1"
    }

    filtro_periodo["grupo"] = filtro_periodo["grupo"].astype(str).str.upper().str.strip()

    for palabra, numero in reemplazos.items():
        filtro_periodo["grupo"] = filtro_periodo["grupo"].str.replace(palabra, numero, regex=False)

    filtro_periodo["grupo"] = filtro_periodo["grupo"].str.replace(".", "", regex=False)
    filtro_periodo["grupo"] = filtro_periodo["grupo"].str.replace(r"\s+", " ", regex=True).str.strip()

    filtro_periodo["nombre"] = (
        filtro_periodo["nombre"]
        .astype(str)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    filtro_periodo["nombre_key"] = filtro_periodo["nombre"].apply(normalizar_texto)

    # ==============================
    # 2. PROCESAR ARCHIVO ACUMULADOS
    # ==============================
    df_acumulados = procesarAcumuladosQ10(archivo_acumulados)

    # ======================
    # 3. HACER EL CRUCE
    # ======================
    resultado = filtro_periodo.merge(
        df_acumulados,
        on="nombre_key",
        how="left",
        suffixes=("_nota", "_ac")
    )

    # Mantener siempre el dato de notas si no hubo match
    resultado["Nombre"] = resultado["nombre_ac"].fillna(resultado["nombre_nota"])
    resultado["GrupoFinal"] = resultado["grupo_ac"].fillna(resultado["grupo_nota"])

    resultado["Promovido"] = resultado["promedio"].apply(lambda x: "Sí" if x >= 3 else "No")
    resultado["Observación"] = resultado["promedio"].apply(
        lambda x: "Promovido" if x >= 3 else "No promovido"
    )

    final = pd.DataFrame({
        "Código de Matrícula": resultado["codigoMatricula"],
        "Folio": resultado["folio"],
        "Nombre": resultado["Nombre"],
        "Tipo Documento": resultado["tipoDocumento"],
        "Num Documento": resultado["numDocumento"],
        "Año Lectivo": resultado["añoLectivo"],
        "Grupo": resultado["GrupoFinal"],
        "Nivel": resultado["nivel"],
        "Promovido": resultado["Promovido"],
        "Observación": resultado["Observación"],
        "PROMEDIO": resultado["promedio"],
        "C.NATURALES Y EDUCACION AMBIENTAL": resultado["ciencias_naturales"],
        "Fisica": resultado["fisica"],
        "Quimica": resultado["quimica"],
        "CIENCIAS POLITICAS Y ECONÓMICAS": resultado["ciencias_politicas_economicas"],
        "CIENCIAS SOCIALES (HIST, GEOG, CONST, DEMOC)": resultado["ciencias_sociales"],
        "CÍVICA Y CONSTITUCIÓN": resultado["civica_y_constitucion"],
        "EDUCACIÓN ARTÍSTICA Y CULTURAL": resultado["educacion_artistica"],
        "EDUCACIÓN CRISTIANA": resultado["educacion_cristiana"],
        "EDUCACIÓN ÉTICA Y VALORES": resultado["educacion_etica"],
        "EDUCACIÓN FÍSICA, RECREACIÓN Y DEPORTES": resultado["educacion_fisica"],
        "filosofia": resultado["filosofia"],
        "HUMANIDADES LENGUA CASTELLANA E IDIOMA EXTRANJERO": resultado["idioma_extranjero"],
        "IDIOMA EXTRANJERO (INGLES)": resultado["idioma_extranjero"],
        "LENGUA CASTELLANA": resultado["lengua_castellana"],
        "MATEMÁTICAS": resultado["matematicas"],
        "TECNOLOGÍA E INFORMÁTICA": resultado["tecnologia"],
    })

    # Si no encontró en familiar, dejar vacíos esos campos
    columnas_vacias = [
        "Código de Matrícula",
        "Folio",
        "Tipo Documento",
        "Num Documento",
        "Año Lectivo",
        "Nivel",
    ]

    for col in columnas_vacias:
        final[col] = final[col].where(final[col].notna(), "")

    final["Nombre"] = final["Nombre"].fillna("")
    final["Grupo"] = final["Grupo"].fillna("")
    final["Promovido"] = final["Promovido"].fillna("")
    final["Observación"] = final["Observación"].fillna("")

    if ruta_salida:
        final.to_excel(ruta_salida, index=False)
        print(f"✅ Archivo Excel creado en: {ruta_salida}")

    return final.reset_index(drop=True)