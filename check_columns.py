import pandas as pd

df = pd.read_excel("EmployeeList.xlsx")
print(df.columns.tolist())