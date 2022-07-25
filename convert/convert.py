import os
import pandas as pd
files = os.listdir()
for f in files:
    if f.endswith(".xls") or f.endswith(".XLS"):
        df = pd.read_excel(f, engine='xlrd')
        writer = pd.ExcelWriter(f[:-4]+".xlsx", engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1')
        writer.save()
