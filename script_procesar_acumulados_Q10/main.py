import pandas as pd
from procesarNotasAcumulados_1_9 import procesarAcumulado_1_9
from procesarNotasAcumulados_10_11 import procesarAcumulado_10_11

resultados = []
AÑO = 2025
data_primeros = procesarAcumulado_1_9('../Datos_base/estudiantes_global_historico _Q10/Sabana_de_notas_descriptiva_1.xlsx', AÑO )
data_segundos = procesarAcumulado_1_9('../Datos_base/estudiantes_global_historico _Q10/Sabana_de_notas_descriptiva_2.xlsx', AÑO )
data_terceros = procesarAcumulado_1_9('../Datos_base/estudiantes_global_historico _Q10/Sabana_de_notas_descriptiva_3.xlsx', AÑO )
data_cuartos = procesarAcumulado_1_9('../Datos_base/estudiantes_global_historico _Q10/Sabana_de_notas_descriptiva_4.xlsx', AÑO )
data_quintos = procesarAcumulado_1_9('../Datos_base/estudiantes_global_historico _Q10/Sabana_de_notas_descriptiva_5.xlsx', AÑO )
data_sextos = procesarAcumulado_1_9('../Datos_base/estudiantes_global_historico _Q10/Sabana_de_notas_descriptiva_6.xlsx', AÑO )
data_septimos = procesarAcumulado_1_9('../Datos_base/estudiantes_global_historico _Q10/Sabana_de_notas_descriptiva_7.xlsx', AÑO )
data_octavos = procesarAcumulado_1_9('../Datos_base/estudiantes_global_historico _Q10/Sabana_de_notas_descriptiva_8.xlsx', AÑO )
data_novenos = procesarAcumulado_1_9('../Datos_base/estudiantes_global_historico _Q10/Sabana_de_notas_descriptiva_9.xlsx', AÑO )
data_decimos = procesarAcumulado_10_11('../Datos_base/estudiantes_global_historico _Q10/Sabana_de_notas_descriptiva_10.xlsx', AÑO )
data_onces = procesarAcumulado_10_11('../Datos_base/estudiantes_global_historico _Q10/Sabana_de_notas_descriptiva_11.xlsx', AÑO )
  

 
consolidado = pd.concat([
    data_primeros,
    data_segundos,
    data_terceros,
    data_cuartos,
    data_quintos,
    data_sextos,
    data_septimos,
    data_octavos,
    data_novenos,
    data_decimos,
    data_onces
    ], ignore_index=True)

ruta_salida = f"./Consolidado de notas del acumulado de {AÑO}.xlsx"
consolidado.to_excel(ruta_salida, index=False)
print("")
print(f"Archivo Excel creado en: {ruta_salida}")
print(f"Filas consolidadoultantes: {len(consolidado)}")
print("")