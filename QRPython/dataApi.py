import requests
import json

# URL de la API
url = "https://backend-m7iv.onrender.com/api/units"

# Realizar la solicitud GET
try:
    response = requests.get(url)
    response.raise_for_status()  # Esto lanzará una excepción si la respuesta contiene un error HTTP
    
    # Parsear la respuesta JSON
    data = response.json()
     

    # Filtrar las unidades que tienen la dirección "C Sicologia"
    unidades_filtradas = []
    
    for unit in data:
        ubicacion = unit['location']
        direccion = ubicacion['direccion']
        
        # Comprobar si la dirección es "C Sicologia"
        if direccion == "C Administración":
            unidades_filtradas.append(unit)

    # Guardar las unidades filtradas en un archivo JSON
    if unidades_filtradas:
        with open("unidades_filtradas_psicologia.json", "w") as json_file:
            json.dump(unidades_filtradas, json_file, indent=4)
        print(f"Se han guardado {len(unidades_filtradas)} registros en el archivo 'unidades_filtradas.json'.")
    else:
        print("No se encontraron unidades con la dirección 'C psicologia'.")
    print(len(unidades_filtradas))
except requests.exceptions.HTTPError as http_err:
    print(f"HTTP error occurred: {http_err}")
except Exception as err:
    print(f"Other error occurred: {err}")
