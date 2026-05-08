import unicodedata
import re
import pandas as pd

def limpiar_nombre(nombre):
    if pd.isna(nombre):
        return ""
    
    nombre = str(nombre).upper().strip()

    # quitar tildes
    nombre = unicodedata.normalize('NFKD', nombre).encode('ascii', 'ignore').decode('utf-8')

    # quitar espacios duplicados
    nombre = re.sub(r'\s+', ' ', nombre)

    return nombre