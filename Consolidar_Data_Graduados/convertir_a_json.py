import pandas as pd
import json
import os
from pymongo import MongoClient
from dotenv import load_dotenv

# ===============================
# CARGAR VARIABLES DE ENTORNO
# ===============================
load_dotenv()

# ===============================
# FUNCIÓN PARA CONVERTIR NÚMEROS
# ===============================
def convert_to_number(value):
    if isinstance(value, str):
        try:
            return float(value.replace(',', '.'))
        except ValueError:
            return value
    return value

try:
    # ===============================
    # 1. CARGAR EXCEL
    # ===============================
    excel_path = './data_certificados.xlsx'
    df = pd.read_excel(
        excel_path,
        sheet_name=0,
        dtype={'Num Documento': str}
    )

    df.columns = df.columns.str.strip()
    print("Columnas detectadas:", df.columns.tolist())

    # ===============================
    # 2. RENOMBRAR COLUMNAS (MAPA)
    # ===============================
    rename_map = {
        'Código de Matrícula': 'codigoMatricula',
        'Folio': 'folio',
        'Nombre': 'nombre',
        'Tipo Documento': 'tipoDocumento',
        'Num Documento': 'numDocumento',
        'Año Lectivo': 'añoLectivo',
        'Grupo': 'grupo',
        'Promovido': 'promovido',
        'Observación': 'observacion',
        'PROMEDIO': 'promedio',

        'C.NATURALES Y EDUCACION AMBIENTAL': 'naturalesYEducacionAmbiental',
        'Fisica': 'Fisica',
        'Quimica': 'Quimica',
        'CIENCIAS POLITICAS Y ECONÓMICAS': 'cienciasPoliticasYEconomicas',
        'CIENCIAS SOCIALES (HIST, GEOG, CONST, DEMOC)': 'cienciasSociales',
        'CÍVICA Y CONSTITUCIÓN': 'civicaYConstitucion',
        'EDUCACIÓN ARTÍSTICA Y CULTURAL': 'educacionArtisticaYCultural',
        'EDUCACIÓN CRISTIANA': 'educacionCristiana',
        'EDUCACIÓN ÉTICA Y VALORES': 'educacionEticaYValores',
        'EDUCACIÓN FÍSICA, RECREACIÓN Y DEPORTES': 'educacionFisicaYRecreacionYDeportes',
        'filosofia': 'filosofia',
        'HUMANIDADES LENGUA CASTELLANA E IDIOMA EXTRANJERO': 'humanidadesLenguaCastellanaEIdiomaExtranjero',
        'IDIOMA EXTRANJERO (INGLES)': 'idiomaExtranjeroIngles',
        'LENGUA CASTELLANA': 'lenguaCastellana',
        'MATEMÁTICAS': 'matematicas',
        'TECNOLOGÍA E INFORMÁTICA': 'tecnologiaEInformatica',

        'dim_socioafectiva': 'dim_socioafectiva',
        'dim_cognitiva': 'dim_cognitiva',
        'dim_etica': 'dim_etica',
        'dim_corporal': 'dim_corporal',
        'dim_comunicativa': 'dim_comunicativa',
        'ingles': 'ingles',
        'dim_espiritual': 'dim_espiritual',
        'dim_estetica': 'dim_estetica',
        'musica': 'musica',
        'dim_acti_valoratica': 'dim_acti_valoratica',

        'RECTOR_A': 'rector',
        'CC_RECTOR_A': 'ccRector',
        'Ciudad_expedición_rector': 'ciudadExpedicionRector',
        'SECRETARIA': 'secretaria',
        'CC_SECRETARIA': 'ccSecretaria',
        'Ciudad_expedición_secretaria': 'ciudadExpedicionSecretaria'
    }

    # ==========================================================
    # 2.5. AGREGAR REGISTRO JSON AL EXCEL (como última fila)
    # ==========================================================
    # ✅ OJO: aquí cargas "este json" (ajusta el nombre/ruta si es otro)
    # ==========================================================
    # 2.5. AGREGAR REGISTROS DEL JSON CONSOLIDADO AL EXCEL
    # ==========================================================
    json_path = './dataCoaweb/consolidado.json'
    if os.path.exists(json_path):
        with open(json_path, 'r', encoding='utf-8') as f:
            nuevo_registro = json.load(f)  # puede ser LISTA o DICCIONARIO

        # Invertir el rename_map: (nombre_final -> nombre_original_excel)
        inverse_map = {v: k for k, v in rename_map.items()}

        # ✅ Asegurar que sea lista
        if isinstance(nuevo_registro, dict):
            registros = [nuevo_registro]
        elif isinstance(nuevo_registro, list):
            registros = nuevo_registro
        else:
            raise ValueError("El JSON no es dict ni list. Formato no soportado.")

        filas_excel = []
        for reg in registros:
            # Convertir llaves del JSON (formato final) a columnas originales del Excel
            reg_excel = {}
            for k, v in reg.items():
                col_original = inverse_map.get(k, k)
                reg_excel[col_original] = v

            filas_excel.append(reg_excel)

        df_nuevos = pd.DataFrame(filas_excel)

        # Asegurar columnas: agregar faltantes en ambos sentidos
        for col in df.columns:
            if col not in df_nuevos.columns:
                df_nuevos[col] = None

        for col in df_nuevos.columns:
            if col not in df.columns:
                df[col] = None

        # Reordenar df_nuevos con el mismo orden de df
        df_nuevos = df_nuevos[df.columns]

        # Concatenar todos los nuevos al final
        df = pd.concat([df, df_nuevos], ignore_index=True)

        # Guardar Excel
        df.to_excel(excel_path, index=False)
        print(f"✅ {len(df_nuevos)} registros del consolidado agregados al final de {excel_path}")
    else:
        print(f"⚠️ No se encontró {json_path}. No se agregó ninguna fila al Excel.")


    # ===============================
    # 3. RENOMBRAR COLUMNAS (aplica a DF ya con la fila nueva)
    # ===============================
    df = df.rename(columns=rename_map)

    # ===============================
    # 4. LIMPIAR TEXTO
    # ===============================
    columnas_a_convertir = [
        'codigoMatricula', 'nombre', 'tipoDocumento', 'numDocumento', 'grupo',
        'promovido', 'observacion', 'rector', 'ccRector', 'secretaria',
        'ccSecretaria', 'ciudadExpedicionRector', 'ciudadExpedicionSecretaria',
        'dim_socioafectiva', 'dim_cognitiva', 'dim_etica', 'dim_corporal',
        'dim_comunicativa', 'ingles', 'dim_espiritual', 'dim_estetica',
        'musica', 'dim_acti_valoratica'
    ]

    for col in columnas_a_convertir:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # ===============================
    # 5. CONVERTIR NÚMEROS
    # ===============================
    columnas_numericas = [
        'promedio',
        'naturalesYEducacionAmbiental',
        'cienciasPoliticasYEconomicas',
        'cienciasSociales',
        'civicaYConstitucion',
        'educacionArtisticaYCultural',
        'educacionCristiana',
        'educacionEticaYValores',
        'educacionFisicaYRecreacionYDeportes',
        'humanidadesLenguaCastellanaEIdiomaExtranjero',
        'idiomaExtranjeroIngles',
        'lenguaCastellana',
        'matematicas',
        'tecnologiaEInformatica',
        'Fisica',
        'Quimica',
        'Filosofia'
    ]

    for col in columnas_numericas:
        if col in df.columns:
            df[col] = df[col].apply(convert_to_number)

    # ===============================
    # 6. GENERAR JSON
    # ===============================
    data = json.loads(df.to_json(orient='records', force_ascii=False))

    with open('archivo_estudiantes.json', 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

    print(f"✅ JSON generado con {len(data)} registros.")

    # ===============================
    # 7. CONEXIÓN A MONGODB (SEGURA)
    # ===============================
    MONGO_URI = os.getenv("MONGO_URI")
    MONGO_DB = os.getenv("MONGO_DB")
    MONGO_COLLECTION = os.getenv("MONGO_COLLECTION")

    if not MONGO_URI:
        raise ValueError("❌ No se encontró MONGO_URI en el archivo .env")

    client = MongoClient(MONGO_URI)
    db = client[MONGO_DB]
    collection = db[MONGO_COLLECTION]

    print("✅ Conectado a MongoDB")

    # ===============================
    # 8. LIMPIAR COLECCIÓN COMPLETA
    # ===============================
    collection.delete_many({})
    print("🧹 Colección limpiada completamente.")

    # ===============================
    # 9. INSERTAR TODO DESDE CERO
    # ===============================
    if data:
        collection.insert_many(data)
        print(f"✅ {len(data)} registros insertados correctamente.")
    else:
        print("⚠️ El JSON está vacío. No se insertó nada.")

except Exception as e:
    print("❌ Error general:", e)
