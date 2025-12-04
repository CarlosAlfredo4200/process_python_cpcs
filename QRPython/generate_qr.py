import requests
import qrcode
from PIL import Image, ImageDraw, ImageFont

# URL de la API
url = "https://backend-m7iv.onrender.com/api/units"

# Cargar una fuente para el texto
def load_font(size):
    try:
        return ImageFont.truetype("arial.ttf", size)
    except IOError:
        return ImageFont.load_default()



# Realizar la solicitud GET
try:
    response = requests.get(url)
    response.raise_for_status()
    
    # Parsear la respuesta JSON
    data = response.json()
    
    # Filtrar las unidades que tienen la dirección "C Sicologia"
    unidades_filtradas = []
    
    for unit in data:
        ubicacion = unit['location']
        direccion = ubicacion['direccion']
        
        # Comprobar si la dirección es "C Sicologia"
        if direccion == "Tesoreria":
            unidades_filtradas.append(unit)


    # Contador para unidades con la dirección "C psicología"
    count_unidades = 0
    
    # Procesar los datos y generar los códigos QR
    for unit in unidades_filtradas:
        # Información del producto
        id_unidad = unit['_id']
        producto = unit['id_producto']
        nombre_producto = producto['name']
        marca = producto['brand']
        modelo = producto['model']
        precio = producto['price']
        imagen = producto['image']['filePath']
        subcategoria = producto['subcategory']
        
        # Información de la ubicación
        ubicacion = unit['location']
        nombre_ubicacion = ubicacion['nombre']
        direccion = ubicacion['direccion']
        estado_ubicacion = ubicacion['estado']
        
        # Estado del producto
        estado_producto = unit['estado']
        
        # Filtrar solo las unidades con la dirección "C psicología"
        if direccion == "Tesoreria":
            count_unidades += 1  # Incrementar el contador
            
            # Generar URL específica para cada unidad usando el id_unidad
            qr_url = f'https://backendqr-ns9q.onrender.com/message/{id_unidad}'
            
            # Crear el código QR
            qr = qrcode.QRCode(version=1, box_size=10, border=4)
            qr.add_data(qr_url)
            qr.make(fit=True)

            # Generar la imagen del código QR
            img = qr.make_image(fill='black', back_color='white')

            # Convertir la imagen QR a una imagen PIL
            img = img.convert("RGB")
            draw = ImageDraw.Draw(img)
            
            # Cargar la fuente
            font = load_font(20)

            # Guardar la imagen final
            img.save(f'{subcategoria}_{id_unidad}.png')

            print(f"Código QR generado con imagen")
            # print(f"Imagen QR guardada como: message_{id_unidad}.png")
    
    # Imprimir el conteo total
    print(f"Total de unidades de {direccion}: {count_unidades}")
    print("=" * 50)

except requests.exceptions.HTTPError as http_err:
    print(f"HTTP error occurred: {http_err}")
except Exception as err:
    print(f"Other error occurred: {err}")
