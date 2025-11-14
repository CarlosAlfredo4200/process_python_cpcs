import json
import pandas as pd
import json
from pymongo import MongoClient


# Leer archivo
archivo = pd.read_excel('../Datos_base/Oportunidades/Oportunidades.xlsx')
df = pd.DataFrame(archivo)

# Limpieza básica
df = df.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x))

df.columns = df.columns.str.strip()


def transformar_nombre(nombre: str) -> str:
    if pd.isna(nombre):
        return ""
    partes = nombre.strip().split()
    if len(partes) < 2:
        return nombre.title()
    if len(partes) >= 4:
        primer_nombre, segundo_nombre, primer_apellido, segundo_apellido = partes[:4]
    elif len(partes) == 3:
        primer_nombre, segundo_nombre, primer_apellido, segundo_apellido = partes[0], "", partes[1], partes[2]
    else:
        primer_nombre, segundo_nombre, primer_apellido, segundo_apellido = partes[0], "", partes[1], ""
    nuevo = " ".join([primer_apellido, segundo_apellido, primer_nombre, segundo_nombre]).strip()
    return nuevo.title()

# Aplicar transformación
df['Estudiantes'] = df['Oportunidad'].apply(transformar_nombre)
df['Madre'] = df['1.4 NOMBRES Y APELLIDOS DE LA MADRE'].apply(transformar_nombre)
df['Padre'] = df['1.1 NOMBRES Y APELLIDOS DEL PADRE'].apply(transformar_nombre)
df['Acudiente_financiero'] = df['1.9 NOMBRE DEL ACUDIENTE FINANCIERO'].apply(transformar_nombre)
df['Acudiente_familiar'] = df['2.0 NOMBRE DEL ACUDIENTE FAMILIAR'].apply(transformar_nombre)


# Renombrar columnas
df = df.rename(columns={
    'Correo electrónico': 'Correo',
    '¿Cómo se enteró?': '¿Cómo_se_enteró?',
    'Medio de contacto': 'Medio_contacto',
    'Identificación':'Identificacion',
    '1.8 DIRECCION DE SU RESIDENCIA': 'Dirección_B',
    '1.3 NUMERO DE TELEFONO O NUMERO DE CELULAR DEL PADRE': 'Contacto_padre',
    '1.7 ESTADO CIVIL DE LOS PADRES': 'Estado_civíl',
    '2.2 NOMBRE DEL DOCENTE QUE ENTREVISTA A LOS PADRES': 'Docente_entrevista_Padres',
    '2.3 NOMBRE DEL DOCENTE QUE ENTREVISTA AL ESTUDIANTE': 'Docente_entrevista_estudiante',
    '2.1 FECHA DEL PROCESO DE ENTREVISTA': 'Fecha_proceso_entrevista',
    '1.5 EMAIL DE LA MADRE': "correo_madre",
    '1.2 EMAIL DEL PADRE': 'Correo_padre',
    '1.6 NUMERO DE TELEFONO O NUMERO DE CELULAR DE LA MADRE': 'Contacto_madre',
    '2.4 CUENTA CON ALGUN DIAGNÓSTICO (FÍSICO, MENTAL, COGNITIVO, ETC)': 'Diagnóstico',
    '2.5 CUENTA CON ALGUN DIAGNÓSTICO ¿ CUAL?': 'Detalle_diagnóstico',
    '2.6 REINGRESO': 'Reingreso',
    '2.7 REQUIERE COMPROMISO DISCIPLINARIO': 'Compromiso_disc',
    '2.8 REQUIERE COMPROMISO PEDAGOGICO': 'Compromiso_pedag',
    '1. FECHA RECIBIDO FORMULARIO': 'Fecha_formulario_recibido',
    'N° FORMULARIO': 'Formulario',
    'AÑO LECTIVO': 'Año_lectivo',
    'Fecha de registro': 'Fecha_Registro',
    'Negocio': 'Negocio',
    'Fecha tentativa de cierre': 'Fecha_tentativa_cierre',
    'Grado': 'Grado',
    '1.1 CONCEPTO DE EDUCACION CRISTIANA': 'Concepto_ed_cristiana',
    '1.2 NOMBRE DE SU IGLESIA O CONGREGACIÓN': 'Congregación',
    '1.3 NOMBRE DE SU PASTOR O LIDER ESPIRITUAL': 'Lider_espiritual',
    '2.0 CONCEPTO DE PSICOLOGIA': 'Concepto_psicologia',
    '2.1 CONCEPTO DE PSICOLOGIA (FAMILIAR)': 'Concepto_psicologia_familia',
    '2.3 CONCEPTO DE PSICOLOGIA (PRUEBAS DE TEST A PADRES)': 'Test_padres',
    '2.4 CONCEPTO DE PSICOLOGIA (PRUEBAS DE TEST A ESTUDIANTES)': 'Test_estudiante',
    '2.2 CONCEPTO DE PSICOLOGIA (ESPIRITUAL)': 'Espiritual',
    '3.0 CONCEPTO DE ADMINISTRACIÓN': 'Concepto_adm',
    '3.1 ¿APRUEBA EL DATACREDITO?': 'Datacredito',
    '3.2 ¿NECESITA CODEUDOR?': 'Codeudor',
    '3.3 OBSERVACIONES DEL ADMINISTRADOR': 'Observaciones_adm',
    '4.0 INGRESO APROBADO POR EL COMITÉ DE ADMISIONES': 'Aprobado_admisiones',
    '4.1 FECHA DE APROBACIÓN POR EL COMITÉ DE ADMISIONES': 'Fecha_aprobado_admisiones',
    '4.2 OBSERVACIONES DE RECTORIA': 'Observaciones_rectoria',
    '4.3 APELACIÓN': 'Apelación',
    '4.4 DETALLES DE LA APELACIÓN': 'Detalles_apelación',
    '4.5 RESULTADOS DE LA APELACIÓN EN SEGUNDA INSTANCIA': 'Resultado_apelación',
    'Estado negocio': 'Estado_negocio',
    'Causa de pérdida': 'Causa_perdida',
})

# Seleccionar columnas finales (sin duplicados)
df = df[[
    'Estudiantes', 'Identificacion', 'Madre', 'Padre',
    'Acudiente_financiero', 'Acudiente_familiar', 'Correo',
    'Celular', 'Teléfono', 'Dirección', 'Municipio', 'Barrio',
    'Asesor', '¿Cómo_se_enteró?', 'Medio_contacto',
    'Contacto_padre', 'Estado_civíl', 'Docente_entrevista_Padres',
    'Docente_entrevista_estudiante', 'Fecha_proceso_entrevista',
    'correo_madre', 'Correo_padre', 'Contacto_madre', 'Diagnóstico',
    'Detalle_diagnóstico', 'Reingreso', 'Compromiso_disc',
    'Compromiso_pedag', 'Fecha_formulario_recibido', 'Formulario',
    'Año_lectivo', 'Fecha_Registro', 'Fecha última actividad',
    'Negocio', 'Fecha_tentativa_cierre', 'Grado',
    'Concepto_ed_cristiana', 'Congregación', 'Lider_espiritual',
    'Concepto_psicologia', 'Concepto_psicologia_familia',
    'Test_padres', 'Test_estudiante', 'Espiritual',
    'Concepto_adm', 'Datacredito', 'Codeudor', 'Observaciones_adm',
    'Aprobado_admisiones', 'Fecha_aprobado_admisiones',
    'Observaciones_rectoria', 'Apelación', 'Detalles_apelación',
    'Resultado_apelación', 'Estado_negocio', 'Causa_perdida'
]]



# Filtrar filas duplicadas (muestra todos menos la primera ocurrencia)
# df_filtrado = df[df['Año_lectivo'].isin([2025, 2026])].copy()
df_filtrado = df[(df['Año_lectivo'] == 2025) | (df['Año_lectivo'] == 2026)].copy()




 

 
# Exportar a JSON normal (lista de objetos)
df_to_json = df_filtrado.to_json(orient='records', lines=False, force_ascii=False)

with open("nuevosProcesados.json", "w", encoding="utf-8") as json_file:
    json.dump(json.loads(df_to_json), json_file, ensure_ascii=False, indent=4)
    

print("Archivo JSON creado exitosamente: nuevosProcesados.json")

# Conexión a MongoDB (ajusta el puerto y la URI si es necesario)
client = MongoClient("mongodb+srv://alfredomontoyacorreo2:dbcpcs2024@cluster0.wu0k7uf.mongodb.net/")

with open("nuevosProcesados.json", "r", encoding="utf-8") as file:
    data = json.load(file)

# Seleccionar la base de datos y colección
db = client["test"]
collection = db["AdmisionesEstudiantes"]
 

collection.delete_many({})  # limpiar todo antes
collection.insert_many(data)
total = collection.count_documents({})
print(f'Insertados en Mongo : {total} registros')


# Exportar a Excel también
ARCHIVO_SALIDA = "./nuevosProcesados.xlsx"
df_filtrado.to_excel(ARCHIVO_SALIDA, index=False)
print(f"Archivo Excel creado exitosamente en: {ARCHIVO_SALIDA}")
