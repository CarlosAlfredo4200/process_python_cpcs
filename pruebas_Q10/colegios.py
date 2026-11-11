
import urllib.request
import urllib.error
import json
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

try:
    # CONFIGURACIÓN API Q10
    url = "https://api.q10.com/v1/instituciones-educativas?limit=1000"

    hdr = {
        "Cache-Control": "no-cache",
        "Accept": "application/json",
        'Api-Key': '23117b580a1048eab7fa33021ad255b4',
    }

    req = urllib.request.Request(
        url,
        headers=hdr,
        method="GET"
    )

    with urllib.request.urlopen(req, timeout=60) as response:
        contenido = response.read().decode("utf-8")
        datos = json.loads(contenido)
        print("Estado HTTP:", response.status)

    if not isinstance(datos, list):
        raise ValueError("La respuesta de Q10 no es una lista.")

    total_registros = len(datos)

    # CREAR ARCHIVO EXCEL
    wb = Workbook()
    ws = wb.active
    ws.title = "Instituciones Q10"

    columnas = [
        ("Consecutivo", "Consecutivo"),
        ("Nombre", "Nombre"),
        ("NIT", "NIT"),
        ("Telefono", "Teléfono"),
        ("Correo_electronico", "Correo electrónico"),
        ("Direccion", "Dirección"),
        ("Codigo_municipio", "Código municipio"),
        ("Nombre_municipio", "Municipio"),
        ("Estado", "Estado")
    ]

    # ENCABEZADOS
    for col, (_, titulo) in enumerate(columnas, start=1):
        celda = ws.cell(row=1, column=col, value=titulo)
        celda.font = Font(color="FFFFFF", bold=True)
        celda.fill = PatternFill(
            start_color="17365D",
            end_color="17365D",
            fill_type="solid"
        )
        celda.alignment = Alignment(horizontal="center")

    # INSERTAR TODOS LOS REGISTROS
    for fila, colegio in enumerate(datos, start=2):

        for col, (campo, _) in enumerate(columnas, start=1):
            valor = colegio.get(campo)

            if valor is None:
                valor = ""

            if campo == "Estado":
                valor = "Activo" if valor is True else (
                    "Inactivo" if valor is False else ""
                )
            else:
                valor = str(valor)

            celda = ws.cell(
                row=fila,
                column=col,
                value=valor
            )

            celda.number_format = "@"

    # ANCHOS DE COLUMNAS
    anchos = [15, 55, 22, 20, 45, 40, 22, 25, 15]

    for indice, ancho in enumerate(anchos, start=1):
        letra = get_column_letter(indice)
        ws.column_dimensions[letra].width = ancho

    # ACTIVAR FILTROS
    ws.auto_filter.ref = ws.dimensions

    # FIJAR PRIMERA FILA
    ws.freeze_panes = "A2"

    # GUARDAR EXCEL
    archivo = "Instituciones_Educativas_Q10.xlsx"
    wb.save(archivo)

    print("\n" + "=" * 60)
    print("ARCHIVO EXCEL GENERADO CORRECTAMENTE")
    print(f"TOTAL DE INSTITUCIONES: {total_registros}")
    print(f"ARCHIVO: {os.path.abspath(archivo)}")
    print("=" * 60)

except urllib.error.HTTPError as e:
    print("Error HTTP:", e.code)
    print(e.read().decode("utf-8", errors="replace"))

except Exception as e:
    print("Error:", e)
