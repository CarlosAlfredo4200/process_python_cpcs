import pandas as pd
import json
from pymongo import MongoClient

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
    df = pd.read_excel('./data_certificados.xlsx', sheet_name=0, dtype={'Num Documento': str})
    df.columns = df.columns.str.strip()

    print("Columnas detectadas:", df.columns.tolist())

    # ===============================
    # 2. RENOMBRAR COLUMNAS
    # ===============================
    df = df.rename(columns={
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
        'Filosofia': 'Filosofia',
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
    })

    # ===============================
    # 3. LIMPIAR TEXTO
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
    # 4. CONVERTIR NÚMEROS
    # ===============================
    columnas_numericas = [
        'promedio', 'naturalesYEducacionAmbiental', 'cienciasPoliticasYEconomicas',
        'cienciasSociales', 'civicaYConstitucion', 'educacionArtisticaYCultural',
        'educacionCristiana', 'educacionEticaYValores',
        'educacionFisicaYRecreacionYDeportes',
        'humanidadesLenguaCastellanaEIdiomaExtranjero',
        'idiomaExtranjeroIngles', 'lenguaCastellana',
        'matematicas', 'tecnologiaEInformatica',
        'Fisica', 'Quimica', 'Filosofia'
    ]

    for col in columnas_numericas:
        if col in df.columns:
            df[col] = df[col].apply(convert_to_number)

    # ===============================
    # 5. GENERAR JSON
    # ===============================
    data = json.loads(df.to_json(orient='records', force_ascii=False))

    with open('archivo_estudiantes.json', 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

    print(f"✅ JSON generado con {len(data)} registros.")

    # ===============================
    # 6. CONEXIÓN A MONGODB
    # ===============================
    client = MongoClient(
        "mongodb+srv://alfredomontoyacorreo2:dbcpcs2024@cluster0.wu0k7uf.mongodb.net/"
    )

    db = client["test"]
    collection = db["studentgraduates"]

    print("✅ Conectado a MongoDB")

    # ===============================
    # 7. LIMPIAR COLECCIÓN COMPLETA
    # ===============================
    collection.delete_many({})
    print("🧹 Colección limpiada completamente.")

    # ===============================
    # 8. INSERTAR TODO DESDE CERO
    # ===============================
    if data:
        collection.insert_many(data)
        print(f"✅ {len(data)} registros insertados correctamente.")
    else:
        print("⚠️ El JSON está vacío. No se insertó nada.")

except Exception as e:
    print("❌ Error general:", e)
