import pandas as pd
import json
data = pd.read_json('./test.asistenciapadres.json')
data_ed = json.dumps(data.to_dict(orient="records"), indent=4)
print(data_ed)