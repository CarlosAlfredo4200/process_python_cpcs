import pandas as pd
import json

def convert_to_number(value):
    """Convertir cadena a número flotante reemplazando comas por puntos. Verifica si el valor es una cadena."""
    if isinstance(value, str):  # Verificar si el valor es una cadena
        try:
            return float(value.replace(',', '.'))
        except ValueError:
            return value  # Devuelve el valor original si no puede convertirlo
    return value  # Devuelve el valor original si no es una cadena

try:
    # Cargar el archivo Excel, asegurando que 'Num Documento' se lea como texto
    df = pd.read_excel('./data_certificados.xlsx', sheet_name=0, dtype={'Num Documento': str})  

    # Limpiar nombres de columnas para eliminar espacios extra
    df.columns = df.columns.str.strip()

    # Verificar los nombres de las columnas
    print("Columnas en el DataFrame después de limpiar:", df.columns.tolist())

    # Renombrar las columnas para que coincidan con la estructura del esquema
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
        'Fisica':"Fisica",
        'Quimica':"Quimica",
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
        
        'dim_socioafectiva':'dim_socioafectiva', 
        'dim_cognitiva' :'dim_cognitiva', 
        'dim_etica' :'dim_etica', 
        'dim_corporal' :'dim_corporal', 
        'dim_comunicativa' :'dim_comunicativa', 
        'ingles' :'ingles', 
        'dim_espiritual' :'dim_espiritual', 
        'dim_estetica' :'dim_estetica', 
        'musica' :'musica', 
        'dim_acti_valoratica' :'dim_acti_valoratica', 
        'RECTOR_A': 'rector',
        'CC_RECTOR_A': 'ccRector',
        'Ciudad_expedición_rector': 'ciudadExpedicionRector',
        'SECRETARIA': 'secretaria',
        'CC_SECRETARIA': 'ccSecretaria',
        'Ciudad_expedición_secretaria': 'ciudadExpedicionSecretaria'
    })

    # Convertir las columnas de texto a formato string y eliminar espacios en blanco
    columnas_a_convertir = [
        'codigoMatricula', 'nombre', 'tipoDocumento', 'numDocumento', 'grupo', 'promovido', 'observacion',
        'rector', 'ccRector', 'secretaria', 'ccSecretaria', 'ciudadExpedicionRector', 'ciudadExpedicionSecretaria', 
        'dim_socioafectiva', 'dim_cognitiva', 'dim_etica','dim_corporal', 'dim_comunicativa', 'ingles', 
        'dim_espiritual', 'dim_estetica', 'musica', 'dim_acti_valoratica'
    ]
    
    for col in columnas_a_convertir:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()  # Asegura que sea string y elimina espacios

    # Convertir las columnas numéricas a números flotantes
    columnas_numericas = [
        'promedio', 'naturalesYEducacionAmbiental', 'cienciasPoliticasYEconomicas', 'cienciasSociales',
        'civicaYConstitucion', 'educacionArtisticaYCultural', 'educacionCristiana', 'educacionEticaYValores',
        'educacionFisicaYRecreacionYDeportes', 'humanidadesLenguaCastellanaEIdiomaExtranjero',
        'idiomaExtranjeroIngles', 'lenguaCastellana', 'matematicas', 'tecnologiaEInformatica',
        'Fisica', 'Quimica', 'Filosofia'
    ]

    for col in columnas_numericas:
        if col in df.columns:
            df[col] = df[col].apply(convert_to_number)

    # Convertir el DataFrame a JSON (sin secuencias de escape)
    data_json = df.to_json(orient='records', lines=False, force_ascii=False)

    # Guardar el archivo JSON en UTF-8
    with open('archivo_estudiantes.json', 'w', encoding='utf-8') as json_file:
        json.dump(json.loads(data_json), json_file, ensure_ascii=False, indent=4)

    print("Archivo JSON guardado exitosamente como 'archivo_estudiantes.json'.")

except pd.errors.ParserError as e:
    print(f"Error al analizar el archivo Excel: {e}")
except Exception as e:
    print(f"Se produjo un error: {e}")
