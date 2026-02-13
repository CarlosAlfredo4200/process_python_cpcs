import os
import re
import json
import glob
import pandas as pd

# =========================
# CONFIG
# =========================
CARPETA_ENTRADA = "./archivo_a_procesar"   # <-- aquí van tus 500 archivos
SALIDA_JSON = "consolidado.json"
SALIDA_ERRORES = "errores_procesamiento.json"

inputColums = [0, 14]  # 0=texto/área, 14=nota

# =========================
# HELPERS
# =========================
def norm_txt(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip())

def norm_upper(s: str) -> str:
    return norm_txt(s).upper()

def to_float_or_nan(v):
    try:
        if v is None or v == "":
            return float("nan")
        return float(v)
    except Exception:
        return float("nan")

def to_float_or_0(v):
    try:
        if v is None or v == "":
            return 0.0
        return round(float(v), 1)
    except Exception:
        return 0.0

def safe_str(v):
    return "" if v is None else str(v).strip()

def convertir_xls_a_xlsx_si_falta(path_xls: str) -> str:
    """
    Convierte a .xlsx si el archivo es .xls y devuelve el path del .xlsx.
    Si ya existe el .xlsx convertido, lo reutiliza.
    """
    base, ext = os.path.splitext(path_xls)
    if ext.lower() != ".xls":
        return path_xls  # ya es xlsx u otro

    path_xlsx = base + "_convertido.xlsx"
    if not os.path.exists(path_xlsx):
        df_temp = pd.read_excel(path_xls)   # lee todo y guarda xlsx
        df_temp.to_excel(path_xlsx, index=False)
    return path_xlsx

# =========================
# MAPEO ÁREAS -> CLAVES
# =========================
AREA_TO_KEY = {
    "C.NATURALES Y EDUCACION AMBIEN": "naturalesYEducacionAmbiental",
    "C. SOC.(HIST,GEOG,CONST,DEMOC)": "cienciasSociales",
    "CIVICA Y CONSTITUCION": "civicaYConstitucion",
    "EDUCACION ARTISTICA Y CULTURAL": "educacionArtisticaYCultural",
    "ED.CRISTIANA": "educacionCristiana",
    "ED.ETICA Y VALORES": "educacionEticaYValores",
    "EDUCACION FISICA, RECREACION Y DEPORTES": "educacionFisicaYRecreacionYDeportes",
    "HUMANIDADES (L. CASTELLANA)": "lenguaCastellana",
    "LENGUA CASTELLANA": "lenguaCastellana",
    "IDIOMA EXTRANJERO (INGLES)": "idiomaExtranjeroIngles",
    "INGLES": "idiomaExtranjeroIngles",
    "MATEMATICAS": "matematicas",
    "TECNOLOGIA E INFORMATICA": "tecnologiaEInformatica",
}

CLAVES_ESPERADAS = [
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

# =========================
# FUNCIÓN: PROCESAR 1 ARCHIVO
# =========================
def procesar_archivo(path_archivo: str) -> dict:
    # 1) Convertir si es .xls
    path_lectura = convertir_xls_a_xlsx_si_falta(path_archivo)

    # 2) Leer columnas requeridas
    dataCoaweb = pd.read_excel(path_lectura, usecols=inputColums)
    dataCoawebED = dataCoaweb.iloc[5:].reset_index(drop=True)

    if len(dataCoawebED) < 2:
        raise ValueError("No hay suficientes filas después del recorte (se requieren al menos 2).")

    fila_1 = dataCoawebED.iloc[0]
    fila_2 = dataCoawebED.iloc[1]

    texto = safe_str(fila_1.iloc[0])
    texto_fila_2 = safe_str(fila_2.iloc[0])

    # 3) Extracciones
    nombre_match = re.search(r"Que\s+(.*?),", texto)
    Nombre = nombre_match.group(1).strip() if nombre_match else ""

    ciudad_match = re.search(r"de\s+(.*?),", texto)
    Ciudad = ciudad_match.group(1).strip() if ciudad_match else ""

    tipoDoc_match = re.search(
    r"identificado\s+con\s+(.*?)\s*N[º°]\.?",
    texto,
    re.IGNORECASE
    )
    tipoDocumento = norm_txt(tipoDoc_match.group(1)) if tipoDoc_match else ""
    tipoDocumento = tipoDocumento.replace(" ", "")  # ✅ "T. I." -> "T.I."

    

    doc_match = re.search(r"N[º°]\.?\s*(\d+)\s*de", texto)
    numDocumento = doc_match.group(1).strip() if doc_match else ""

    nivel_match = re.search(r"grado\s+[\w\-]+\s+de\s+(.*?)\.", texto, re.IGNORECASE)
    Nivel = norm_txt(nivel_match.group(1)) if nivel_match else ""

    # Grado: "04-B" -> "4 B"
    grado_match = re.search(r"grado\s+([\w\-]+)\s+de", texto, re.IGNORECASE)
    Grado_raw = grado_match.group(1).strip() if grado_match else ""
    # robusto: acepta "04-B", "04 B", "04B"
    m = re.match(r"^\s*(\d+)\s*[-\s]?\s*([A-Za-z])\s*$", Grado_raw)
    if m:
        Grado = f"{int(m.group(1))} {m.group(2).upper()}"
    else:
        Grado = Grado_raw

    anio_match = re.search(r"año\s+lectivo\s+(\d{4})", texto, re.IGNORECASE)
    Anio = anio_match.group(1).strip() if anio_match else ""

    matricula_match = re.search(r"Matricula.*?N[º°]\.?\s*([0-9]+)\s*y", texto_fila_2, re.IGNORECASE)
    Matricula = int(matricula_match.group(1).strip()) if matricula_match else ""

    folio_match = re.search(r"folio.*?N[º°]\.?\s*([0-9]+)", texto_fila_2, re.IGNORECASE)
    Folio = folio_match.group(1).strip() if folio_match else ""

    # 4) Tabla áreas/notas
    dataColumnas = dataCoawebED.iloc[2:200].reset_index(drop=True)
    COL_AREA = dataColumnas.columns[0]
    COL_NOTA = dataColumnas.columns[1]

    df_notas = dataColumnas.copy()

    # Quitar header
    if len(df_notas) > 0 and norm_upper(df_notas.iloc[0][COL_AREA]) in ["AREA/MATERIA", "AREA", "MATERIA"]:
        df_notas = df_notas.iloc[1:].reset_index(drop=True)

    df_notas[COL_AREA] = df_notas[COL_AREA].astype(str).map(norm_txt)
    df_notas[COL_NOTA] = pd.to_numeric(df_notas[COL_NOTA], errors="coerce")
    df_notas = df_notas[df_notas[COL_AREA].ne("")].reset_index(drop=True)

    # 5) Mapear notas
    notas_por_clave = {}
    for _, row in df_notas.iterrows():
        area_raw = safe_str(row[COL_AREA])
        nota = row[COL_NOTA]
        area_norm = norm_upper(area_raw)

        clave = None
        for k, v in AREA_TO_KEY.items():
            if norm_upper(k) == area_norm:
                clave = v
                break

        if clave:
            notas_por_clave[clave] = to_float_or_0(nota)

    # 6) Completar faltantes
    for k in CLAVES_ESPERADAS:
        if k not in notas_por_clave:
            notas_por_clave[k] = 4  # (tu default)

    # 7) JSON final (incluye el archivo origen para rastrear)
    resultado = {
        "sourceFile": os.path.basename(path_archivo),

        "codigoMatricula": to_float_or_nan(Matricula),
        "folio": to_float_or_nan(Folio),
        "nombre": Nombre,
        "ciudad": Ciudad,
        "tipoDocumento": tipoDocumento,
        "numDocumento": numDocumento,
        "añoLectivo": to_float_or_nan(Anio),
        "grupo": Grado,
        "Nivel": Nivel,
        "promovido": "Sí",
        "observacion": "Promovido",
        "promedio": 4.6,
        **notas_por_clave,

        "dim_socioafectiva": "nan",
        "dim_cognitiva": "nan",
        "dim_etica": "nan",
        "dim_corporal": "nan",
        "dim_comunicativa": "nan",
        "ingles": "nan",
        "dim_espiritual": "nan",
        "dim_estetica": "nan",
        "musica": "nan",
        "dim_acti_valoratica": "nan",

        "rector": "EDGAR GONZALEZ FUENTES",
        "ccRector": "16678768.0",
        "ciudadExpedicionRector": "CALI",
        "secretaria": "LUZ MIRIYAM BOTERO ZABALA",
        "ccSecretaria": "43208497.0",
        "ciudadExpedicionSecretaria": "MEDELLÍN",
    }

    return resultado

# =========================
# MAIN: RECORRER TODOS LOS ARCHIVOS
# =========================
def main():
    # toma .xls y .xlsx
    archivos = sorted(glob.glob(os.path.join(CARPETA_ENTRADA, "*.xls"))) + \
               sorted(glob.glob(os.path.join(CARPETA_ENTRADA, "*.xlsx")))

    if not archivos:
        raise FileNotFoundError(f"No encontré .xls o .xlsx en: {CARPETA_ENTRADA}")

    consolidado = []
    errores = []

    for i, path in enumerate(archivos, start=1):
        try:
            res = procesar_archivo(path)
            consolidado.append(res)
            print(f"✅ [{i}/{len(archivos)}] OK: {os.path.basename(path)}")
        except Exception as e:
            print(f"❌ [{i}/{len(archivos)}] ERROR: {os.path.basename(path)} -> {e}")
            errores.append({"file": os.path.basename(path), "error": str(e)})

    with open(SALIDA_JSON, "w", encoding="utf-8") as f:
        json.dump(consolidado, f, ensure_ascii=False, indent=2)

    with open(SALIDA_ERRORES, "w", encoding="utf-8") as f:
        json.dump(errores, f, ensure_ascii=False, indent=2)

    print(f"\n📦 Consolidado: {SALIDA_JSON} (registros={len(consolidado)})")
    print(f"🧯 Errores: {SALIDA_ERRORES} (errores={len(errores)})")

if __name__ == "__main__":
    main()
