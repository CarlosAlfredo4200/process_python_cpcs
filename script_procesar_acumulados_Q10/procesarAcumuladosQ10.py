import pandas as pd
import re
import unicodedata


def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    valor = str(valor).strip().upper()
    valor = unicodedata.normalize("NFKD", valor).encode("ascii", "ignore").decode("utf-8")
    valor = re.sub(r"\s+", " ", valor)
    return valor


def procesarAcumuladosQ10(
    archivo="../Datos_base/estudiantes_global_historico _Q10/Estudiantes__Informacion_familiar.xlsx"
):
    input_columns = [
        'Código de matrícula',
        'Primer nombre',
        'Segundo nombre',
        'Primer apellido',
        'Segundo apellido',
        'Tipo de identificación',
        'Número de identificación',
        'Folio',
        'Grupo',
        'Grado',
        'Año lectivo',
        'Nivel'
    ]

    data = pd.read_excel(archivo, usecols=input_columns)
    df = pd.DataFrame(data)

    # Limpiar espacios
    df[input_columns] = df[input_columns].apply(
        lambda col: col.str.strip() if col.dtype == "object" else col
    )

    # Construir nombre
    df['nombre'] = (
        df['Primer apellido'].fillna('') + ' ' +
        df['Segundo apellido'].fillna('') + ' ' +
        df['Primer nombre'].fillna('') + ' ' +
        df['Segundo nombre'].fillna('')
    ).str.replace(r'\s+', ' ', regex=True).str.strip()

    # Renombrar columnas
    df = df.rename(columns={
        'Código de matrícula': 'codigoMatricula',
        'Tipo de identificación': 'tipoDocumento',
        'Número de identificación': 'numDocumento',
        'Folio': 'folio',
        'Grupo': 'grupo',
        'Año lectivo': 'añoLectivo',
        'Nivel': 'nivel'
    })

    # Reemplazar vacíos
    df = df.replace(r'^\s*$', pd.NA, regex=True)

    # Eliminar filas completamente vacías en campos importantes
    df = df.dropna(subset=['codigoMatricula', 'folio', 'numDocumento', 'nombre'], how='all')

    # Normalizar grupo
    df['grupo'] = df['grupo'].replace({
        'PRIMERO A': '1 A',
        'SEGUNDO A': '2 A',
        'TERCERO A': '3 A',
        'CUARTO A': '4 A',
        'QUINTO A': '5 A',
        'SEXTO A': '6 A',
        'SEPTIMO A': '7 A',
        'OCTAVO A': '8 A',
        'NOVENO A': '9 A',
        'DECIMO A': '10 A',
        'ONCE A': '11 A',
        'PRIMERO B': '1 B',
        'SEGUNDO B': '2 B',
        'TERCERO B': '3 B',
        'CUARTO B': '4 B',
        'QUINTO B': '5 B',
        'SEXTO B': '6 B',
        'SEPTIMO B': '7 B',
        'OCTAVO B': '8 B',
        'NOVENO B': '9 B',
        'DECIMO B': '10 B',
        'DECIMO B1': '10 B1',
        'ONCE B': '11 B',
        'PRIMERO C': '1 C',
        'SEGUNDO C': '2 C',
        'TERCERO C': '3 C',
        'CUARTO C': '4 C',
        'QUINTO C': '5 C',
        'SEXTO C': '6 C',
        'SEPTIMO C': '7 C',
        'OCTAVO C': '8 C',
        'NOVENO C': '9 C',
        'DECIMO C': '10 C',
        'ONCE C': '11 C',
    })

    # Normalizar nivel
    df['nivel'] = df['nivel'].replace({
        'PREESCOLAR': 'Preescolar',
        'BÁSICA PRIMARIA': 'Básica Primaria',
        'BÁSICA SECUNDARIA': 'Básica Secundaria',
        'MEDIA ACADÉMICA': 'Media Académica',
    })

    # Ajustar año lectivo
    df['añoLectivo'] = df['añoLectivo'].replace({
        '2026 PRE ESCOLAR': '2026',
        '2025 PRE ESCOLAR': '2025',
        '2024 PRE ESCOLAR': '2024',
        '2023 PRE ESCOLAR': '2023',
        '2022 PRE ESCOLAR': '2022',
        '2021 PRE ESCOLAR': '2021',
        '2020 PRE ESCOLAR': '2020',
    })

    # Tipo documento
    df['tipoDocumento'] = df['tipoDocumento'].replace({
        'Registro Civil de Nacimiento': 'RCN',
        'Tarjeta de Identidad': 'T.I.',
        'Cédula de Ciudadanía': 'C.C.',
        'Cédula ó Identificación de Extranjería': 'C.E',
        'Permiso por Protección Temporal': 'Otro',
    })

    # Convertir números
    df['codigoMatricula'] = pd.to_numeric(df['codigoMatricula'], errors='coerce').astype('Int64')
    df['folio'] = pd.to_numeric(df['folio'], errors='coerce').astype('Int64')

    # Clave para cruce
    df['nombre_key'] = df['nombre'].apply(normalizar_texto)

    # Limpiar duplicados por nombre normalizado
    df = df.drop_duplicates(subset=['nombre_key'], keep='first')

    columnas_finales = [
        'codigoMatricula',
        'folio',
        'nombre',
        'nombre_key',
        'tipoDocumento',
        'numDocumento',
        'añoLectivo',
        'grupo',
        'nivel'
    ]

    return df[columnas_finales].reset_index(drop=True)