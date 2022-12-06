import os
import pandas as pd
files = os.listdir()
for f in files:
    """
    if f.endswith(".xls") or f.endswith(".XLS"):
        df = pd.read_excel(f, engine="xlrd")
        writer = pd.ExcelWriter(f[:-4]+".xlsx", engine="xlsxwriter")
        df.to_excel(writer, sheet_name="Sheet1", index=False, header=False)
        writer.save()
    """
    if f.endswith('.xlsx') or f.endswith('.XLSX'):
        df = pd.read_excel(f, engine="openpyxl")
        writer = pd.ExcelWriter(f[:-5]+".xls", engine="openpyxl")
        df.to_excel(writer, sheet_name="Sheet1", index=False, header=False)
        writer.save()

