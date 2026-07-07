import requests
import json
from dotenv import load_dotenv
import os

load_dotenv()

URL = os.getenv("API_URL")

response = requests.get(URL)

if response.status_code == 200:
    data = response.json()

    with open("inventario.json", "w", encoding="utf-8") as archivo:
        json.dump(data, archivo, ensure_ascii=False, indent=4)

    print("JSON guardado correctamente")
else:
    print(f"Error: {response.status_code}")