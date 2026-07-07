import pandas as pd
import json

with open('./Datos_2026/asistencias_ep_ (1).json', "r", encoding="utf-8") as f:
    data = json.load(f)

df = pd.json_normalize(data)

def contar_asistidas(lista_asistencias):
    if not isinstance(lista_asistencias, list):
        return 0

    return sum(1 for a in lista_asistencias if a.get("asistio") == True)

def calcular_asistencias(lista_asistencias):
    if not isinstance(lista_asistencias, list):
        return "0 de 0"

    total = len(lista_asistencias)
    asistidas = contar_asistidas(lista_asistencias)

    return f"{asistidas} de {total}"

def calcular_certificado(row):
    escuela = row["escuela.nombre"]
    asistidas = contar_asistidas(row["asistencias"])

    if escuela == "HABLANDO DE SEXUALIDAD EN CASA" and asistidas >= 3:
        return True

    if escuela == "CIBERFAMILIAS 3D" and asistidas >= 4:
        return True

    if escuela == "EL ARTE DE SER PADRES PRESENCIAL" and asistidas >= 4:
        return True

    return False

df["asistencias_resumen"] = df["asistencias"].apply(calcular_asistencias)

df["certificadoOtorgado"] = df.apply(calcular_certificado, axis=1)

df["año"] = "2026-A"

resultado_excel = df[
    [
        "estudiante.nombre",
        "estudiante.documento",
        "estudiante.grupo",
        "escuela.nombre",
        "asistencias_resumen",
        "certificadoOtorgado",
        "año"
    ]
]

resultado_excel.to_excel(
    "reporte_asistencias.xlsx",
    index=False
)

resultado_json = []

for _, row in resultado_excel.iterrows():
    resultado_json.append({
        "estudiante": {
            "nombre": row["estudiante.nombre"],
            "documento": str(row["estudiante.documento"]),
            "grupo": row["estudiante.grupo"]
        },
        "escuela": {
            "nombre": row["escuela.nombre"]
        },
        "asistencias_resumen": row["asistencias_resumen"],
        "certificadoOtorgado": bool(row["certificadoOtorgado"]),
        "año": row["año"]
    })

with open("reporte_asistencias.json", "w", encoding="utf-8") as f:
    json.dump(
        resultado_json,
        f,
        ensure_ascii=False,
        indent=4
    )

print("Archivo Excel y JSON generados correctamente")