import pandas as pd
import re

inputColums = [0, 11, 14, 16]
dataCoaweb = pd.read_excel('./informe-pacsis (2).xls', usecols=inputColums)

dataCoawebED = dataCoaweb.iloc[5:].reset_index(drop=True)

# Obtener solo la primera fila
fila = dataCoawebED.iloc[0]

# Obtener texto de la primera columna
texto = str(fila.iloc[0])

print("Texto completo:", texto)

# Extraer nombre
nombre_match = re.search(r'Que\s+(.*?),', texto)
nombre = nombre_match.group(1).strip() if nombre_match else ""

# Extraer la ciudad
ciudad_match = re.search(r'de\s+(.*?),',texto)
ciudad = ciudad_match.group(1).strip() if ciudad_match else ""

# Extraer documento
doc_match = re.search(r'N[º°]\.?\s*(\d+)\s*de', texto)
documento = doc_match.group(1) if doc_match else ""

print("Nombre:", nombre)
print("Documento:", documento)
print("Ciudad:", ciudad)

