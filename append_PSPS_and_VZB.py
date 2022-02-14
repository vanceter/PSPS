# 2022/02/11 Terry Vance
# importing the modules needed
import pandas as pd
import xlsxwriter

df = pd.read_excel("PSPS_MAIN.xlsx")
df2 = pd.read_excel("2022_PSPS_VZB_Sites.xlsx")

df.append(df2, ignore_index=True)
df.to_excel("PSPS_MAIN_VZB.xlsx", index = False, sheet_name='PSPS_MAIN_VZB')

# df = pd.read_excel('tmp.xlsx', sheet_name=None, index_col=None)
# cdf = pd.concat(df.values())

#print(df)