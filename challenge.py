from IPython.display import display
import pandas as pd

table = pd.read_excel('teste_ar2.xlsx', sheet_name=1, usecols='A:F', skiprows=6)


print(table.values)