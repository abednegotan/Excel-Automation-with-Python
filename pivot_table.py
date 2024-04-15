import pandas as pd

df = pd.read_excel('superstore.xls')

df = df[['Ship Mode', 'Category', 'Sales']]

print(df)

pivot_table = df.pivot_table(index='Ship Mode', columns = 'Category', values = 'Sales', aggfunc = 'sum').round(0)

print(pivot_table)

pivot_table.to_excel('pivot_table.xlsx', 'Report', startrow = 4)