import pandas as pd  

inputColumns =[0,11]
dataBase = pd.read_excel('./dataHistoricaPacsis/informe-pacsis.xls', usecols=inputColumns)

    # --- Elimina las primeras 8 filas ---
dataEstudiante = dataBase.iloc[6:].reset_index(drop=True)
print(dataEstudiante.head())