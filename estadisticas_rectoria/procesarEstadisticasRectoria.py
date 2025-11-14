import pandas as pd
import matplotlib.pyplot as plt
import plotly.express as px


# Leer el archivo JSON
consolidadoPeriodos = pd.read_json('../script_procesarNotas_final/consolidado.json')

# --- Eliminar columnas específicas ---
columnas_a_eliminar = [
    "codigo",
    "nombre",
    "puesto",
    "promedio",
    "observaciones_ciencias_naturales",         
    "observaciones_fisica",                      
    "observaciones_quimica",                      
    "observaciones_ciencias_politicas_economicas",
    "observaciones_ciencias_sociales",            
    "observaciones_civica_y_constitucion",        
    "observaciones_educacion_artistica",          
    "observaciones_educacion_cristiana",          
    "observaciones_educacion_etica",              
    "observaciones_educacion_fisica",             
    "observaciones_filosofia",                    
    "observaciones_idioma_extranjero",            
    "observaciones_lengua_castellana",            
    "observaciones_matematicas",                  
    "observaciones_tecnologia",                   
    "metas_ciencias_naturales",                   
    "metas_fisica",                               
    "metas_quimica",                              
    "metas_ciencias_politicas_economicas",        
    "metas_ciencias_sociales",                    
    "metas_civica_y_constitucion",                
    "metas_educacion_artistica",                  
    "metas_educacion_cristiana",                  
    "metas_educacion_etica",                      
    "metas_educacion_fisica",                     
    "metas_filosofia",                            
    "metas_idioma_extranjero",                    
    "metas_lengua_castellana",                    
    "metas_matematicas",                          
    "metas_tecnologia",                           
    "rep_eva_ciencias_naturales",                 
    "rep_eva_fisica",                             
    "rep_eva_quimica",                            
    "rep_eva_ciencias_politicas_economicas",      
    "rep_eva_ciencias_sociales",                  
    "rep_eva_civica_y_constitucion",              
    "rep_eva_educacion_artistica",                
    "rep_eva_educacion_cristiana",                
    "rep_eva_educacion_etica",                    
    "rep_eva_educacion_fisica",                   
    "rep_eva_filosofia",                          
    "rep_eva_idioma_extranjero",                  
    "rep_eva_lengua_castellana",                  
    "rep_eva_matematicas",                        
    "rep_eva_tecnologia"
]

consolidadoPeriodos.drop(columnas_a_eliminar, axis=1, inplace=True)
grupos = consolidadoPeriodos["grupo"]
 
# --- Filtrar datos por periodo ---
periodo = "PERIODO 2"
grupo = "1. A "
periodo2 = consolidadoPeriodos[consolidadoPeriodos['periodo'] == periodo].copy()


materias = [
"ciencias_naturales",
"fisica",                       
 "quimica",                      
 "ciencias_politicas_economicas",
 "ciencias_sociales",            
 "civica_y_constitucion",        
 "educacion_artistica",          
 "educacion_cristiana",          
 "educacion_etica",              
 "educacion_fisica",             
 "filosofia",                    
 "idioma_extranjero",            
 "lengua_castellana",            
 "matematicas",                  
 "tecnologia",
]



# promedios = consolidadoPeriodos[materias].mean().round(1)
# Agrupar por grupo y calcular el promedio de cada materia
promedios_por_grupo = periodo2.groupby("grupo")[materias].mean().round(1)  # redondeamos a 2 decimales

promedios_por_grupo = promedios_por_grupo.rename(columns={
"ciencias_naturales":"C.Naturales",
"fisica":"Física",                       
 "quimica":"Química",                      
 "ciencias_politicas_economicas": "C.Poli y Econ",
 "ciencias_sociales": "C.Sociales",            
 "civica_y_constitucion": "Civica y Const",        
 "educacion_artistica": "E.Artistica",          
 "educacion_cristiana":"E.Cristiana",          
 "educacion_etica": "E.Etica",              
 "educacion_fisica": "E.Fisica",             
 "filosofia": "Filosofía",                    
 "idioma_extranjero": "I.Extranjero",            
 "lengua_castellana": "L.Castellana",            
 "matematicas": "Matematicas",                  
 "tecnologia": "Tecnología"
})




# promedios_por_grupo
# # Exportar a Excel
# ruta_salida = "./promedios_periodo_2.xlsx"
# promedios_por_grupo.to_excel(ruta_salida, index=True)
# print(f"Archivo Excel creado en: {ruta_salida}")
 