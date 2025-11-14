import pandas as pd
from promedios_funcion import conectar_mongo, calcular_promedios_por_grupo

def main():
    collection = conectar_mongo()
    data = list(collection.find())
    df = pd.DataFrame(data)

    # Definir columnas y limpiar texto
    areas = [
        'lengua_castellana',
        'matematicas',
        'ciencias_naturales',
        'ciencias_sociales', 
        'idioma_extranjero' 
    ]
    orden_columnas = ['grupo','codigo','nombre','periodo'] + areas
    df = df[orden_columnas]
    df["grupo"] = df["grupo"].str.strip()
    df["periodo"] = df["periodo"].str.strip().str.upper()

    # Lista de grupos
    grupos = ["11. A", "11. B", "10. A", "10. B", "9. A", "9. B","9. B1", "8. A", "8. B","7. A", "7. B", "6. A", "6. B", "5. A", "5. B", "4. A", "4. B","4. C", "3. A", "3. B", "2. A", "2. B","2. C","1. A", "1. B",]

    resultados_globales = []

    # Calcular promedios por cada grupo
    for grupo in grupos:
        print(f"📊 Procesando grupo: {grupo}")
        df_prom = calcular_promedios_por_grupo(df, grupo)
        resultados_globales.append(df_prom)

    # Unir todos los resultados
    df_final = pd.concat(resultados_globales, ignore_index=True)

    df_final = df_final.rename(columns={
         'ciencias_naturales':'CN',
         'fisica':'FI',
         'quimica':'QU',
         'ciencias_politicas_economicas':'CP',
        'ciencias_sociales':'CS',
        'civica_y_constitucion':'CYC',
        'educacion_artistica':'ART',
        'educacion_cristiana':'EC',
        'educacion_etica':'EE',
        'educacion_fisica':'EF',
        'filosofia':'FL',
        'idioma_extranjero':'ING',
        'lengua_castellana':'LC',
        'matematicas':'MAT',
        'tecnologia':'TEC'
        
    })
    # 🔹 Redondear a un decimal
    df_final = df_final.round(1)

    # Exportar a Excel
    ruta_salida = "./promedios_todos_grupos.xlsx"
    df_final.to_excel(ruta_salida, index=False)
    print(f"\n✅ Archivo Excel creado correctamente en: {ruta_salida}")
    # print(df_final)


if __name__ == "__main__":
    main()
