import pandas as pd
import json

def procesar_certificados(archivo_excel, hoja):
    data_historico = pd.read_excel(archivo_excel, sheet_name=hoja)

    # Limpiar strings
    for col in data_historico.select_dtypes(include="object").columns:
        data_historico[col] = data_historico[col].astype(str).str.strip()

    # Convertir números
    data_historico["conceptoNum"] = pd.to_numeric(
        data_historico["conceptoNum"], errors="coerce"
    )

    resultado = []

    for identificacion, grupo_df in data_historico.groupby("nombre"):
        primera = grupo_df.iloc[0]

        estudiante = {
            "folio": str(primera["folio"]),
            "nombre": str(primera["nombre"]),
            "tipoDocumento": str(primera["tipoDocumento"]),
            "numDocumento": str(primera["numDocumento"]),
            "Grado": str(primera["grupo"]),
            "Nivel": str(primera["seccion"]),
            "añoLectivo": str(primera["añoLectivo"]),
            "codigoMatricula": str(primera["codigoMatricula"]),
            "grupo": str(primera["grupo"]),
        }

        for _, fila in grupo_df.iterrows():
            area = str(fila["NombreArea"]).strip()
            valor = fila["conceptoNum"]

            if pd.notna(valor):
                estudiante[area] = float(valor)

        resultado.append(estudiante)

    # Crear DataFrame
    df_resultado = pd.DataFrame(resultado)

    # Agregar columnas nuevas
    df_resultado["promovido"] = "Sí"
    df_resultado["observacion"] = "Promovido"
    df_resultado["promedio"] = 4.1
    df_resultado["Fisica"] = 0
    df_resultado["Quimica"] = 0
    df_resultado["humanidadesLenguaCastellanaEIdiomaExtranjero"] = 0

    # Orden personalizado de áreas
    orden_areas = [
        "naturalesYEducacionAmbiental",
        "Fisica",
        "Quimica",
        "cienciasPoliticasYEconomicas",
        "cienciasSociales",
        "civicaYConstitucion",
        "educacionArtisticaYCultural",
        "educacionCristiana",
        "educacionEticaYValores",
        "educacionFisicaYRecreacionYDeportes",
        "filosofia",
        "humanidadesLenguaCastellanaEIdiomaExtranjero",
        "idiomaExtranjeroIngles",
        "lenguaCastellana",
        "matematicas",
        "tecnologiaEInformatica",
    ]

    # Columnas principales
    orden_columnas = [
        "codigoMatricula", "folio", "nombre", "tipoDocumento",
        "numDocumento", "añoLectivo", "grupo", "Nivel",
        "promovido", "observacion", "promedio"
    ]

    # Solo áreas presentes
    areas_presentes = [col for col in orden_areas if col in df_resultado.columns]

    # Otras columnas no definidas
    otras_columnas = [
        col for col in df_resultado.columns
        if col not in orden_columnas + areas_presentes
    ]

    # Orden final
    df_resultado = df_resultado[orden_columnas + areas_presentes + otras_columnas]

    # Nombre de salida según hoja
    nombre_base = str(hoja).replace(" ", "_")

    # Guardar JSON
    with open(f"resultado_{nombre_base}.json", "w", encoding="utf-8") as f:
        json.dump(
            df_resultado.to_dict(orient="records"),
            f,
            ensure_ascii=False,
            indent=4
        )

    # Guardar Excel
    df_resultado.to_excel(f"resultado_{nombre_base}.xlsx", index=False)

    print(f"✅ Procesado correcto de la hoja: {hoja}")
    print(f"✅ JSON generado: resultado_{nombre_base}.json")
    print(f"✅ Excel generado: resultado_{nombre_base}.xlsx")


# ===== USO =====
procesar_certificados("./DATA_CERTIFICADOS_VALORATIVOS.xlsx", "2003")