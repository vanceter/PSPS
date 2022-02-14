# 2022/02 Terry Vance (vanceter)
# OpsTracker file merge for PSPS events.
# importing the module
import pandas as pd
import xlsxwriter

# NOTE each of the underlying xls files needs to have the PSLC value - in some of them, the column header needs to be renamed from PS Loc
# Also need to make sure you export OpsTracker files with the file name option checked
 
# reading only the columns needed from each file
# documentation on pandas read_excel https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html
f_sites = pd.read_excel("opstracker_sites.xlsx", usecols=['SITE_NAME','ADDRESS','CITY','COUNTY','PSLC','POWER_METER', 'GEN_STATUS','GEN_PORTABLE_PLUG', 'GEN_PORTABLE_PLUG_TYPE', 'IS_HUB','IS_HUB_MICROWAVE','REMOTE_MONITORING','SITETECH_NAME','SITETECH_MANAGER_NAME', 'POWER_COMPANY'])
f_gens = pd.read_excel("opstracker_generators.xlsx", usecols=['PSLC', 'FUEL_TYPE1'])
f_cells = pd.read_excel("NorCal_CellInfo.xlsx", usecols=['PSLC', 'eNodeB'])
f_cells5g = pd.read_excel("norcal_cell_info_5g.xlsx", usecols=['PSLC', 'GNODEB'])
f_pge = pd.read_excel("PSPS_FIRE_TIER.xlsx", usecols=['PSLC', 'Fire Tier', 'PSPS PROB', 'PG&E Fee Property'])
f_vzb = pd.read_excel("PSPS_VZB_Sites.xlsx")

# merging the files using PSLC as the index. There are some duplicates in gen and sites files, lots of duplicates in the cell info because of B2B and 5G gNodeBs
# documentation on pandas merge https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.merge.html?highlight=merge#pandas.DataFrame.merge
f_merged = f_sites.merge(f_gens, left_on="PSLC", right_on="PSLC", how="left")
f_merged = f_merged.merge(f_pge, left_on="PSLC", right_on="PSLC", how="left")
f_merged = f_merged.merge(f_cells, left_on="PSLC", right_on="PSLC", how="left")
f_merged = f_merged.merge(f_cells5g, left_on="PSLC", right_on="PSLC", how="left")

# same function but only combining the sites, gens and PSPS/PGE files for Ops
f_merged_ops = f_sites.merge(f_gens, left_on="PSLC", right_on="PSLC", how="left")
f_merged_ops = f_merged_ops.merge(f_pge, left_on="PSLC", right_on="PSLC", how="left")
# create a new dataframe to concatenate the merged data with a static VZB/VZS file of meter numbers from PGE for non-wireless locations
# requires that the PSPS_VZB_Sites.xlsx file exist in the directory, same format as PSPS_MAIN, but with the random meters PGE provides for non-VZW locations in scope
# uses pd.concat https://pandas.pydata.org/pandas-docs/stable/user_guide/merging.html
frames = [f_merged_ops, f_vzb]
concat_ops = pd.concat(frames)
frames_sp = [f_merged, f_vzb]
concat_sp = pd.concat(frames_sp)

# https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.replace.html

# Bulk replace 0 for No and 1 for yes in the various columns
concat_ops['REMOTE_MONITORING'] = concat_ops['REMOTE_MONITORING'].map({0:'No', 1:'Yes',"VZB FACILITY":"VZB FACILITY", "VZW RETAIL SALES":"VZW RETAIL SALES"})
concat_ops['IS_HUB_MICROWAVE'] = concat_ops['IS_HUB_MICROWAVE'].map({0:'No', 1:'Yes',"VZB FACILITY":"VZB FACILITY", "VZW RETAIL SALES":"VZW RETAIL SALES"})
concat_ops['IS_HUB'] = concat_ops['IS_HUB'].map({0:'No', 1:'Yes',"VZB FACILITY":"VZB FACILITY", "VZW RETAIL SALES":"VZW RETAIL SALES"})
concat_sp['REMOTE_MONITORING'] = concat_sp['REMOTE_MONITORING'].map({0:'No', 1:'Yes',"VZB FACILITY":"VZB FACILITY", "VZW RETAIL SALES":"VZW RETAIL SALES"})
concat_sp['IS_HUB_MICROWAVE'] = concat_sp['IS_HUB_MICROWAVE'].map({0:'No', 1:'Yes',"VZB FACILITY":"VZB FACILITY", "VZW RETAIL SALES":"VZW RETAIL SALES"})
concat_sp['IS_HUB'] = concat_sp['IS_HUB'].map({0:'No', 1:'Yes',"VZB FACILITY":"VZB FACILITY", "VZW RETAIL SALES":"VZW RETAIL SALES"})

# creating 2 new files, the PSPS_Main for Gennie, and a version of it with eNB/gNB for SP
# Format the PSPS_MAIN sheet for Ops
# establish the xlsxwriter functionality, defining "writer" as the variable for the workbook filename
writer = pd.ExcelWriter('PSPS_MAIN.xlsx', engine='xlsxwriter')
# Create the merged sheet and output to the file name based on the writer variable
#f_merged_ops.to_excel(writer, index=False, sheet_name='PSPS_MAIN',columns=['POWER_METER','Fire Tier', 'PSPS PROB','PSLC', 'PG&E Fee Property', 'SITE_NAME', 'ADDRESS','CITY','COUNTY', 'GEN_STATUS','FUEL_TYPE1', 'GEN_PORTABLE_PLUG', 'GEN_PORTABLE_PLUG_TYPE', 'REMOTE_MONITORING', 'IS_HUB_MICROWAVE', 'IS_HUB','SITETECH_NAME','SITETECH_MANAGER_NAME', 'POWER_COMPANY'])
concat_ops.to_excel(writer, index=False, sheet_name='PSPS_MAIN',columns=['POWER_METER','Fire Tier', 'PSPS PROB','PSLC', 'PG&E Fee Property', 'SITE_NAME', 'ADDRESS','CITY','COUNTY', 'GEN_STATUS','FUEL_TYPE1', 'GEN_PORTABLE_PLUG', 'GEN_PORTABLE_PLUG_TYPE', 'REMOTE_MONITORING', 'IS_HUB_MICROWAVE', 'IS_HUB','SITETECH_NAME','SITETECH_MANAGER_NAME', 'POWER_COMPANY'])

# Establish the workbook variable
workbook = writer.book

# Setup some formating definitions
# formatting for any cells/columns that need to be center justified
cell_format_center = workbook.add_format()
cell_format_center.set_align('center')

# Define the worksheet variable
worksheet = writer.sheets['PSPS_MAIN']
# Apply some formatting to groups of columns, including cell width and applying the cell formatting previously defined as appropriate
worksheet.set_column('A:A', 20, cell_format_center)
worksheet.set_column('B:D', 10, cell_format_center)
worksheet.set_column('E:E', 18, cell_format_center)
worksheet.set_column('F:G', 44)
worksheet.set_column('H:I', 22)
worksheet.set_column('J:P', 22, cell_format_center)
worksheet.set_column('Q:Q', 22)
worksheet.set_column('R:R', 28)
worksheet.set_column('S:S', 28, cell_format_center)
# Set some worksheet formatting, including creating filter dropdowns and freeze the top row
worksheet.freeze_panes(1, 0)
worksheet.autofilter('A1:S9999')
# Save the sheet
writer.save()

# Format the PSPS_MAIN_SP sheet for SP
# establish the xlsxwriter functionality, defining "writer" as the variable for the workbook filename
writer_sp = pd.ExcelWriter('PSPS_MAIN_SP.xlsx', engine='xlsxwriter')
# Create the merged sheet and output to the file name based on the writer variable
concat_sp.to_excel(writer_sp, index = False, sheet_name='PSPS_MAIN_SP', columns=['POWER_METER','Fire Tier', 'PSPS PROB','PSLC', 'PG&E Fee Property', 'SITE_NAME', 'ADDRESS','CITY','COUNTY', 'GEN_STATUS','FUEL_TYPE1', 'GEN_PORTABLE_PLUG', 'GEN_PORTABLE_PLUG_TYPE', 'REMOTE_MONITORING', 'IS_HUB_MICROWAVE', 'IS_HUB','SITETECH_NAME','SITETECH_MANAGER_NAME', 'POWER_COMPANY', 'eNodeB', 'GNODEB'])

# Establish the workbook variable
workbook_sp = writer_sp.book

# Setup some formating definitions
# formatting for any cells/columns that need to be center justified
cell_format_center_sp = workbook_sp.add_format()
cell_format_center_sp.set_align('center')

# Define the worksheet variable
worksheet_sp = writer_sp.sheets['PSPS_MAIN_SP']
# Apply some formatting to groups of columns, including cell width and applying the cell formatting previously defined as appropriate
worksheet_sp.set_column('A:A', 20, cell_format_center_sp)
worksheet_sp.set_column('B:D', 10, cell_format_center_sp)
worksheet_sp.set_column('E:E', 18, cell_format_center_sp)
worksheet_sp.set_column('F:G', 44)
worksheet_sp.set_column('H:I', 22)
worksheet_sp.set_column('J:P', 22, cell_format_center_sp)
worksheet_sp.set_column('Q:Q', 22)
worksheet_sp.set_column('R:R', 28)
worksheet_sp.set_column('S:S', 28, cell_format_center_sp)
worksheet_sp.set_column('T:U', 10, cell_format_center_sp)
# Set some worksheet formatting, including creating filter dropdowns and freeze the top row
worksheet_sp.freeze_panes(1, 0)
worksheet_sp.autofilter('A1:U9999')
# Save the sheet
writer_sp.save()