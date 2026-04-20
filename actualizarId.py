
import json
import re
import unicodedata
from pathlib import Path

import pandas as pd

# ==============================
# RUTAS
# ==============================
EXCEL_PATH = r"./nombre_facturas.xlsx"
ESTUDIANTES_JSON_PATH = r"./test.estudianterecaudos.json"
FACTURAS_JSON_PATH = r"./test.facturas.json"
OUTPUT_JSON_PATH = r"/mnt/data/facturas_actualizadas.json"
REPORTE_PATH = r"/mnt/data/reporte_actualizacion_facturas.xlsx"

# ==============================
# HELPERS
# ==============================
def normalizar_texto(texto):
    """
    Normaliza texto para comparar nombres:
    - quita tildes
    - mayúsculas
    - espacios duplicados
    """
    if texto is None:
        return ""
    texto = str(texto).strip()
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    texto = texto.upper()
    texto = re.sub(r"\s+", " ", texto)
    return texto


def cargar_json(path_json):
    with open(path_json, "r", encoding="utf-8") as f:
        return json.load(f)


def guardar_json(path_json, data):
    with open(path_json, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ==============================
# CARGA DE DATOS
# ==============================
df_excel = pd.read_excel(EXCEL_PATH)

# Ajusta estos nombres si cambia el Excel
COL_FACTURA = "N° Factura"
COL_NOMBRE = "Nombre Estudiante"

facturas = cargar_json(FACTURAS_JSON_PATH)
estudiantes = cargar_json(ESTUDIANTES_JSON_PATH)

# ==============================
# ÍNDICES
# ==============================
# Índice de estudiantes por nombre normalizado
estudiantes_por_nombre = {}
duplicados_estudiantes = {}

for est in estudiantes:
    nombre_norm = normalizar_texto(est.get("nombre", ""))
    oid_est = est.get("_id", {}).get("$oid")

    if not nombre_norm or not oid_est:
        continue

    if nombre_norm in estudiantes_por_nombre:
        duplicados_estudiantes.setdefault(nombre_norm, [estudiantes_por_nombre[nombre_norm]])
        duplicados_estudiantes[nombre_norm].append(est)
    else:
        estudiantes_por_nombre[nombre_norm] = est

# Índice de facturas por numero_factura
facturas_por_numero = {}
for fac in facturas:
    numero = str(fac.get("numero_factura", "")).strip()
    if numero:
        facturas_por_numero[numero] = fac

# ==============================
# ACTUALIZACIÓN
# ==============================
actualizadas = 0
sin_estudiante = []
sin_factura = []
filas_ok = []

for _, row in df_excel.iterrows():
    numero_factura = str(row.get(COL_FACTURA, "")).strip()
    nombre_excel = str(row.get(COL_NOMBRE, "")).strip()
    nombre_norm = normalizar_texto(nombre_excel)

    if not numero_factura or not nombre_excel:
        continue

    estudiante = estudiantes_por_nombre.get(nombre_norm)
    factura = facturas_por_numero.get(numero_factura)

    if estudiante is None:
        sin_estudiante.append({
            "numero_factura": numero_factura,
            "nombre_excel": nombre_excel,
            "motivo": "Nombre no encontrado en estudiantes"
        })
        continue

    if factura is None:
        sin_factura.append({
            "numero_factura": numero_factura,
            "nombre_excel": nombre_excel,
            "estudiante_oid_encontrado": estudiante["_id"]["$oid"],
            "motivo": "Factura no encontrada en facturas"
        })
        continue

    oid_anterior = factura.get("estudianteId", {}).get("$oid")
    oid_nuevo = estudiante["_id"]["$oid"]

    factura["estudianteId"] = {"$oid": oid_nuevo}

    filas_ok.append({
        "numero_factura": numero_factura,
        "nombre_excel": nombre_excel,
        "nombre_estudiante_json": estudiante.get("nombre", ""),
        "oid_anterior": oid_anterior,
        "oid_nuevo": oid_nuevo,
        "cambio": "SI" if oid_anterior != oid_nuevo else "NO"
    })

    if oid_anterior != oid_nuevo:
        actualizadas += 1

# ==============================
# GUARDAR RESULTADOS
# ==============================
guardar_json(OUTPUT_JSON_PATH, facturas)

with pd.ExcelWriter(REPORTE_PATH, engine="openpyxl") as writer:
    pd.DataFrame(filas_ok).to_excel(writer, index=False, sheet_name="actualizadas_y_revisadas")
    pd.DataFrame(sin_estudiante).to_excel(writer, index=False, sheet_name="sin_estudiante")
    pd.DataFrame(sin_factura).to_excel(writer, index=False, sheet_name="sin_factura")

# ==============================
# RESUMEN
# ==============================
print("Proceso finalizado.")
print(f"Facturas revisadas desde Excel: {len(df_excel)}")
print(f"Facturas con estudianteId cambiado: {actualizadas}")
print(f"Nombres no encontrados en estudiantes: {len(sin_estudiante)}")
print(f"Facturas no encontradas en JSON de facturas: {len(sin_factura)}")
print(f"JSON actualizado: {OUTPUT_JSON_PATH}")
print(f"Reporte Excel: {REPORTE_PATH}")

if duplicados_estudiantes:
    print("\n⚠️ Hay nombres duplicados en estudiantes. Revísalos manualmente:")
    for nombre, registros in duplicados_estudiantes.items():
        print(f"- {nombre}: {len(registros) + 1} registros")
