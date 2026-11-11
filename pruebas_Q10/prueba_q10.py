
import urllib.request
import urllib.error
import json
import os

try:
    url = "https://api.q10.com/v1/codeudores"

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

    with urllib.request.urlopen(req, timeout=30) as response:

        print("Estado HTTP:", response.getcode())

        # Leer la respuesta una sola vez
        contenido = response.read().decode("utf-8")

        # Convertir texto a JSON
        datos = json.loads(contenido)
        # Contar registros
        total_registros = len(datos)

        

        # Organizar JSON con sangría de 2 espacios
        json_ordenado = json.dumps(
            datos,
            indent=2,
            ensure_ascii=False
        )

        # Mostrar JSON organizado
        print(json_ordenado)

        # Guardar respuesta en un archivo
        with open("codeudores.json", "w", encoding="utf-8") as archivo:
            archivo.write(json_ordenado)

        
        print("\nArchivo codeudores.json guardado correctamente.")

except urllib.error.HTTPError as e:
    print("Error HTTP:", e.code)
    print(e.read().decode("utf-8", errors="replace"))
    print(f"\nTotal de registros encontrados: {total_registros}")

except json.JSONDecodeError:
    print("La respuesta no contiene un JSON válido.")

except KeyError:
    print("Falta configurar la variable Q10_API_KEY.")

except Exception as e:
    print("Error:", e)
