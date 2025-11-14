# generar_contrasenas_excel.py
import secrets
import string
from openpyxl import Workbook

def generate_password(length: int = 21) -> str:
    """Genera una contraseña segura de 'length' caracteres."""
    upper = string.ascii_uppercase
    lower = string.ascii_lowercase
    digits = string.digits
    symbols = "!@#$%^&*()-_=+[]{};:,.<>?/~`"

    # Asegurar al menos un carácter de cada tipo
    password = [
        secrets.choice(upper),
        secrets.choice(lower),
        secrets.choice(digits),
        secrets.choice(symbols)
    ]

    pool = upper + lower + digits + symbols
    for _ in range(length - 4):
        password.append(secrets.choice(pool))

    secrets.SystemRandom().shuffle(password)
    return ''.join(password)

def generar_para_correos_excel(correos, archivo_excel="contrasenas.xlsx"):
    # Crear libro y hoja
    wb = Workbook()
    ws = wb.active
    ws.title = "Contraseñas"

    # Encabezados
    ws.append(["Correo", "Contraseña"])

    # Llenar datos
    for correo in correos:
        contrasena = generate_password(21)
        ws.append([correo, contrasena])

    # Guardar en archivo
    wb.save(archivo_excel)
    print(f"✅ Archivo Excel generado: {archivo_excel}")


if __name__ == "__main__":
    correos = [
        "academico@midominio.edu",
        "rectoria@midominio.edu",
        "coordinacion@midominio.edu",
        "secretaria@midominio.edu"
    ]

    generar_para_correos_excel(correos)
