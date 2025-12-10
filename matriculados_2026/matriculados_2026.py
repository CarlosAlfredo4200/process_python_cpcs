import pandas as pd

# Columnas a usar
input_columns = [0,2,5,7,16,18]
data = pd.read_excel('./data/3eb0d84e7ef447a4a7b0a5f0402fb1be.xlsx', usecols=input_columns)

# --- Elimina primeras filas no útiles ---
data_ed = data.iloc[16:].reset_index(drop=True)

# Renombrar columnas
data_ed = data_ed.rename(columns={
    'Unnamed: 0':'Estudiante',
    'Unnamed: 2':'Acudiente',
    'Unnamed: 5':'Grado',
    'Unnamed: 7':'Concepto',
    'Unnamed: 16':'Abono',
    'Unnamed: 18':'Estado',
})

# Limpiar nombres
data_ed["Estudiante"] = data_ed["Estudiante"].astype(str).str.split("\n").str[0].str.strip()
data_ed["Acudiente"]  = data_ed["Acudiente"].astype(str).str.split("\n").str[0].str.strip()

# Extraer Año
data_ed["Año"] = data_ed["Grado"].astype(str).str.extract(r'(20\d{2})', expand=False)

# Extraer solo el grado limpio
data_ed["Grado_Limpio"] = data_ed["Grado"].astype(str).str.extract(r'([A-Z\-]+)', expand=False)
data_ed["Grado"] = data_ed["Grado_Limpio"]
data_ed.drop(columns=["Grado_Limpio"], inplace=True)

# Rellenar hacia abajo
data_ed = data_ed.ffill()

# Quitar filas basura
data_ed = data_ed[~data_ed["Estudiante"].str.contains("Total Principal|www.q10.com", na=False)]

# Convertir 'nan' textual a NaN real y eliminar vacíos
data_ed["Estudiante"] = data_ed["Estudiante"].replace(["nan","None",""], pd.NA)
data_ed = data_ed.dropna(subset=["Estudiante"])

# FILTRAR SOLO AÑO 2026
data_ed = data_ed[data_ed["Año"] == "2026"]

# PAGADOS BASE COMPLETA
data_ed_PAGADOS = data_ed[data_ed["Estado"] == "PAGADO"]

# PAGADOS POR GRADO
data_ed_PAGADOS_PRE_JARDIN = data_ed_PAGADOS[data_ed_PAGADOS["Grado"]=="PRE-JARDIN"]
data_ed_PAGADOS_JARDIN     = data_ed_PAGADOS[data_ed_PAGADOS["Grado"]=="JARDIN"]
data_ed_PAGADOS_TRANSICI   = data_ed_PAGADOS[data_ed_PAGADOS["Grado"]=="TRANSICION"]
data_ed_PAGADOS_PRIMERO    = data_ed_PAGADOS[data_ed_PAGADOS["Grado"]=="PRIMERO"]
data_ed_PAGADOS_SEGUNDO    = data_ed_PAGADOS[data_ed_PAGADOS["Grado"]=="SEGUNDO"]
data_ed_PAGADOS_TERCERO    = data_ed_PAGADOS[data_ed_PAGADOS["Grado"]=="TERCERO"]
data_ed_PAGADOS_CUARTO     = data_ed_PAGADOS[data_ed_PAGADOS["Grado"]=="CUARTO"]
data_ed_PAGADOS_QUINTO     = data_ed_PAGADOS[data_ed_PAGADOS["Grado"]=="QUINTO"]
data_ed_PAGADOS_SEXTO      = data_ed_PAGADOS[data_ed_PAGADOS["Grado"]=="SEXTO"]
data_ed_PAGADOS_SEPTIMO    = data_ed_PAGADOS[data_ed_PAGADOS["Grado"]=="SEPTIMO"]
data_ed_PAGADOS_OCTAVO     = data_ed_PAGADOS[data_ed_PAGADOS["Grado"]=="OCTAVO"]
data_ed_PAGADOS_NOVENO     = data_ed_PAGADOS[data_ed_PAGADOS["Grado"]=="NOVENO"]
data_ed_PAGADOS_DECIMO     = data_ed_PAGADOS[data_ed_PAGADOS["Grado"]=="DECIMO"]
data_ed_PAGADOS_ONCE       = data_ed_PAGADOS[data_ed_PAGADOS["Grado"]=="ONCE"]

# --------------------------------------------------------------------
# EXPORTAR BASE
# --------------------------------------------------------------------
ruta_salida = "./base_matriculados_2026.xlsx"
data_ed.to_excel(ruta_salida, index=False)
print("\n✔ Archivo generado:", ruta_salida)

# --------------------------------------------------------------------
# 🔥 ESTILOS Y RESUMEN FINAL EN EXCEL
# --------------------------------------------------------------------
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

wb = load_workbook(ruta_salida)
ws = wb.active

# Encabezados estilo
header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
header_font = Font(color="FFFFFF", bold=True, size=12)
for cell in ws[1]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center")

# Auto ancho de columnas
for column in ws.columns:
    max_len = 0
    col = column[0].column_letter
    for cell in column:
        if cell.value:
            max_len = max(max_len, len(str(cell.value)))
    ws.column_dimensions[col].width = max_len + 2

# Crear fila donde inicia el resumen
fila = ws.max_row + 2

totales = {
    "TOTAL GENERAL": len(data_ed),
    "TOTAL PAGADOS": len(data_ed_PAGADOS),
    "PAGADOS PRE-JARDIN": len(data_ed_PAGADOS_PRE_JARDIN),
    "PAGADOS JARDIN": len(data_ed_PAGADOS_JARDIN),
    "PAGADOS TRANSICION": len(data_ed_PAGADOS_TRANSICI),
    "PAGADOS PRIMERO": len(data_ed_PAGADOS_PRIMERO),
    "PAGADOS SEGUNDO": len(data_ed_PAGADOS_SEGUNDO),
    "PAGADOS TERCERO": len(data_ed_PAGADOS_TERCERO),
    "PAGADOS CUARTO": len(data_ed_PAGADOS_CUARTO),
    "PAGADOS QUINTO": len(data_ed_PAGADOS_QUINTO),
    "PAGADOS SEXTO": len(data_ed_PAGADOS_SEXTO),
    "PAGADOS SEPTIMO": len(data_ed_PAGADOS_SEPTIMO),
    "PAGADOS OCTAVO": len(data_ed_PAGADOS_OCTAVO),
    "PAGADOS NOVENO": len(data_ed_PAGADOS_NOVENO),
    "PAGADOS DECIMO": len(data_ed_PAGADOS_DECIMO),
    "PAGADOS ONCE": len(data_ed_PAGADOS_ONCE),
}

for key, value in totales.items():
    ws[f"A{fila}"] = key
    ws[f"A{fila}"].font = Font(bold=True, size=11)
    ws[f"B{fila}"] = value
    ws[f"B{fila}"].font = Font(bold=True, size=12, color="007300")
    fila += 1

# Nota final
nota = "Nota: 18 son deudores morosos con más de 4 pensiones. 12 son de fundación y 20 sin reserva de cupo."
ws[f"A{fila+1}"] = nota
ws[f"A{fila+1}"].font = Font(bold=True, italic=True, size=11, color="000080")
ws.merge_cells(f"A{fila+1}:E{fila+1}")
ws[f"A{fila+1}"].alignment = Alignment(horizontal="left")

wb.save(ruta_salida)
print("\n📄 Nota final agregada con éxito.\n")
