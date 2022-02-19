# 2022/02/11 Terry Vance
# importing the modules needed
import pandas as pd
import xlsxwriter

df = pd.read_excel("PSPS_MAIN.xlsx")
df2 = pd.read_excel("PSPS_VZB_Sites.xlsx")
frames = [df, df2]
df3 = pd.concat(frames)
#print(df.concat)
writer_vzb = pd.ExcelWriter('PSPS_MAIN_VZB.xlsx', engine='xlsxwriter')
df3.to_excel(writer_vzb, index=False, sheet_name='PSPS_MAIN_VZB')

writer_vzb.save()