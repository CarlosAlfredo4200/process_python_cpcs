from dbfread import DBF
import pandas as pd

archivo_dbf = r"./C012001/"

tabla = DBF(archivo_dbf, encoding='latin-1')

df = pd.DataFrame(iter(tabla))


print(df.head())