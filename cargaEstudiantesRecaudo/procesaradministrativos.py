import pandas as pd 
import json

data_administrativos = pd.read_excel('./Administrativos.xlsx')
data_docentes = pd.read_excel('./Docentes.xlsx')
 
df = pd.concat([data_administrativos,data_docentes], ignore_index=True)


df['nombre'] = (
    df['Primer nombre'].astype(str).str.strip() + ' ' +
    df['Segundo nombre'].astype(str).str.strip() + ' ' +
    df['Primer apellido'].astype(str).str.strip() + ' ' +
    df['Segundo apellido'].astype(str).str.strip()
)

df['grado'] = 'Empleado'

df = df.rename(columns={
    'Número de identificación':'documentoIdentidad'
})

df['documentoIdentidad'] = pd.to_numeric(
    df['documentoIdentidad'], errors='coerce'
).astype('Int64')

# ===============================
# 8. ELIMINAR REGISTROS VACÍOS
# ===============================
df = df[df['documentoIdentidad'].notna()]

ordenar = ['nombre','documentoIdentidad', 'grado']
df = df[ordenar]


 # 6. GENERAR JSON
    # ===============================
df = json.loads(df.to_json(orient='records', force_ascii=False))

with open('administrativosYdocentes.json', 'w', encoding='utf-8') as f:
    json.dump(df, f, ensure_ascii=False, indent=4)

print(f"✅ JSON generado con {len(df)} registros.")
 