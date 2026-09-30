import pandas as pd
from pathlib import Path

arq = Path("Planinha teste V 1.xlsx")
print("Existe:", arq.exists(), "|", arq.resolve())
try:
    df = pd.read_excel(arq, sheet_name=None)
    for aba, d in df.items():
        print(aba, d.shape)
except Exception as e:
    print(type(e).__name__, "->", e)