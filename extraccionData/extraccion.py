import dbfread
import pandas as pd

table = dbfread.DBF("./factpen.dbf", encoding="latin1")
df = pd.DataFrame(iter(table))
df.to_excel("datos_extraidos.xlsx", index=False)
