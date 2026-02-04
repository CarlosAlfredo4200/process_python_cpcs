import pandas as pd 

input_columns = [1,3]
data_docentes = pd.read_excel('./data/Control_de_acceso_docentes_y_administrativos (1).xlsx', usecols=input_columns)
print(data_docentes.head(10))