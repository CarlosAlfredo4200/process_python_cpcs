import pandas as pd
from procesar_oportunidades import procesarOportunidadesFN
from procesar_solicitudes import procesarSolicitudesFN
from procesar_cartera_pensiones import procesarCarteraFN
import json
from pymongo import MongoClient

resultOportunidades = procesarOportunidadesFN('./data/Oportunidades.xlsx')
print("")
resultSolicitudes = procesarSolicitudesFN('./data/Solicitudes.xlsx')
print("")
# resultCartera = procesarCarteraFN('./data/cd45c78992754f7484ab16f9540d20d1.xlsx')