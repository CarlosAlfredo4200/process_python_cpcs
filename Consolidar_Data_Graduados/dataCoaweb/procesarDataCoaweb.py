import os
import re
import pandas as pd

# ---------------- Config ----------------
archivo_xls = "./informe-pacsis (2).xls"
archivo_xlsx = "./informe-pacsis_convertido.xlsx"
inputColums = [0, 11, 14, 16]

# --------- Convertir SOLO una vez ---------
# (El WARNING solo puede salir la primera vez, porque hay que leer el .xls para convertirlo)
if not os.path.exists(archivo_xlsx):
    df_temp = pd.read_excel(archivo_xls)
    df_temp.to_excel(archivo_xlsx, index=False)
    print("✅ Archivo convertido a .xlsx:", archivo_xlsx)

# --------- Leer siempre el .xlsx ---------
dataCoaweb = pd.read_excel(archivo_xlsx, usecols=inputColums)
dataCoawebED = dataCoaweb.iloc[5:].reset_index(drop=True)

# Validación mínima
if len(dataCoawebED) < 2:
    raise ValueError("El archivo no tiene suficientes filas después del recorte (se requieren al menos 2).")

fila_1 = dataCoawebED.iloc[0]
fila_2 = dataCoawebED.iloc[1]

texto = str(fila_1.iloc[0])
texto_fila_2 = str(fila_2.iloc[0])

# ---------------- Extracciones ----------------

# Nombre
nombre_match = re.search(r"Que\s+(.*?),", texto)
Nombre = nombre_match.group(1).strip() if nombre_match else ""

# Ciudad (primera ocurrencia de "de ...," que suele ser la ciudad del documento)
ciudad_match = re.search(r"de\s+(.*?),", texto)
Ciudad = ciudad_match.group(1).strip() if ciudad_match else ""

# Tipo Documento
tipoDoc_match = re.search(r"con\s+(.*?).", texto)
tipoDocumento = tipoDoc_match.group(1).strip() if tipoDoc_match else ""

# Documento
doc_match = re.search(r"N[º°]\.?\s*(\d+)\s*de", texto)
numDocumento = doc_match.group(1).strip() if doc_match else ""

# Nivel (Básica Primaria, Básica Secundaria, etc.)
nivel_match = re.search(r"grado\s+[\w\-]+\s+de\s+(.*?)\.", texto)
Nivel = nivel_match.group(1).strip() if nivel_match else ""

# Grado (04-B, 1B, etc.)
grado_match = re.search(r"grado\s+([\w\-]+)\s+de", texto, re.IGNORECASE)
Grado = grado_match.group(1).strip() if grado_match else ""

# Año lectivo
anio_match = re.search(r"año\s+lectivo\s+(\d{4})", texto, re.IGNORECASE)
Anio = anio_match.group(1).strip() if anio_match else ""

# Segunda fila: Matrícula
matricula_match = re.search(r"Matricula.*?N[º°]\.?\s*(\d+)\s*y", texto_fila_2, re.IGNORECASE)
Matricula = matricula_match.group(1).strip() if matricula_match else ""

# Segunda fila: Folio
folio_match = re.search(r"folio.*?N[º°]\.?\s*(\d+)", texto_fila_2, re.IGNORECASE)
Folio = folio_match.group(1).strip() if folio_match else ""


dataColumnas = dataCoawebED.iloc[2:15].reset_index(drop=True)

print("Tipo de documento :", tipoDocumento)

