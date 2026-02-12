import pandas as pd
import re

inputColums = [0, 11, 14, 16]
dataCoaweb = pd.read_excel('./informe-pacsis (2).xls', usecols=inputColums)

dataCoawebED = dataCoaweb.iloc[5:].reset_index(drop=True)
print(dataCoawebED.head(14))

# Texto de la fila 0, columna 0
texto = str(dataCoaweb.iloc[0, 0])

# Extraer nombre: después de "Que " hasta la primera coma
nombre_match = re.search(r'Que\s+(.*?),', texto)
nombre = nombre_match.group(1).strip() if nombre_match else ""

# Extraer documento: después de "Nº." hasta "de"
doc_match = re.search(r'N[º°]\.?\s*(\d+)\s*de', texto)
documento = doc_match.group(1) if doc_match else ""

print("Nombre:", nombre)
print("Documento:", documento)
