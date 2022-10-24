from pathlib import Path

import pandas as pd

path = Path(r"C:\Users\bulat\AppData\Roaming\Microsoft\Excel\dataTest309917513254617196\dataTest((Autorecovered-309917763879236613)).xlsb")
df = pd.read_excel(path, engine='pyxlsb', sheet_name="Рез_Муж")
df.to_excel(r'D:\work\ParserExcelData\result.xlsx')