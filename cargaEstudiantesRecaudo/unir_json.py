import json

# ===============================
# 1. LEER ARCHIVOS
# ===============================
with open('./estudiantes.json', 'r', encoding='utf-8') as f:
    data1 = json.load(f)

with open('./administrativosYdocentes.json', 'r', encoding='utf-8') as f:
    data2 = json.load(f)

# ===============================
# 2. UNIR
# ===============================
data_unida = data1 + data2

# ===============================
# 3. ELIMINAR DUPLICADOS
# ===============================
vistos = set()
data_limpia = []

for item in data_unida:
    doc = item.get('documentoIdentidad')

    # validar que exista y no sea null
    if doc is None:
        continue

    # evitar duplicados
    if doc not in vistos:
        vistos.add(doc)
        data_limpia.append(item)

# ===============================
# 4. GUARDAR
# ===============================
with open('unido.json', 'w', encoding='utf-8') as f:
    json.dump(data_limpia, f, ensure_ascii=False, indent=4)

print(f"✅ JSON unido sin duplicados: {len(data_limpia)} registros")