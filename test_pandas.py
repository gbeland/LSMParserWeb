
import pandas as pd
try:
    df = pd.read_excel("TestLog.xlsx", engine="openpyxl", header=None)
    print("Pandas read success")
    print(df.head())
except Exception as e:
    print(f"Pandas failed: {e}")
