# # 2023/10 Terry Vance (vanceter)
# OpsTracker file merge for SCE PSPS events.
# Copy values for the block codes. Then remove the vlookup columns and the extra tabs.
# importing the module
import pandas as pd
import xlsxwriter

file_path = "/Users/txvance/Documents/PSPS/"
file_path_raw = "/Users/txvance/Documents/PSPS/OpsTracker_Raw_Files/"

# NOTE each of the underlying xls files needs to have the PSLC value - in some of them, the column header needs to be renamed from PS Loc
# Also need to make sure you export OpsTracker files with the file name option checked, and change the filter to get -all- sites
# Raw data files need to be here: /Users/txvance/Documents/PSPS/OpsTracker_Raw_Files/
# Output files will go here: /Users/txvance/Documents/PSPS/Tracker/
# 2023/10 Notes for what to do after the file is created - run macro SCE_Format_PSPS in excel: 
# macro saved in personal macro on macbook as: PSPS_Format().vb


 
# reading only the columns needed from each file
# documentation on pandas read_excel https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html
f_sites = pd.read_excel("/Users/txvance/Documents/PSPS/OpsTracker_Raw_Files/opstracker_sites.xlsx", usecols=['SITE_NAME','ADDRESS','CITY','COUNTY','PSLC','POWER_METER', 'GEN_PORTABLE_PLUG', 'GEN_PORTABLE_PLUG_TYPE', 'GO95_FIRE_ZONE_SECTOR', 'SITE_STATUS', 'IS_HUB','IS_HUB_MICROWAVE','REMOTE_MONITORING','SITETECH_NAME','SITEMGR_NAME', 'POWER_COMPANY', 'MDG_ID'])
f_gens = pd.read_excel("/Users/txvance/Documents/PSPS/OpsTracker_Raw_Files/opstracker_generators.xlsx", usecols=['PSLC', 'GEN_STATUS', 'SERIALNUM', 'FUEL_TANK1', 'FUEL_TYPE1', 'MANUFACTURER', 'MODEL', 'GEN_SIZE'])
#f_gens = pd.read_excel("/Users/txvance/Documents/PSPS/OpsTracker_Raw_Files/Asset_Generator_WS_Full_Data_data.xlsx", usecols=['PSLC_CODE', 'STATUS', 'SERIAL_NUMBER', 'GENERATOR_SIZE', 'FUEL_TYPE (CMPL_GENERATOR_SPEC)'])

# Rename column names for Gennie, using OT file
g={'GEN_STATUS':'GEN Y/N','FUEL_TANK1':'TANK SIZE','GEN_SIZE':'GEN SIZE','FUEL_TYPE1':'FUEL TYPE'}
# New rename for Fuze gen version
s={'SITE_NAME':'SITE NAME','GO95_FIRE_ZONE_SECTOR':'FIRE TIER', 'GEN_PORTABLE_PLUG':'PLUG Y/N','GEN_PORTABLE_PLUG_TYPE':'PLUG TYPE','REMOTE_MONITORING':'RM Y/N','IS_HUB':'HUB Y/N', 'IS_HUB_MICROWAVE':'M/W HUB Y/N','SITETECH_NAME':'FIELD ENGINEER','SITEMGR_NAME':'OPS MANAGER','POWER_COMPANY':'POWER COMPANY','POWER_METER':'POWER METER','SITE_STATUS':'SITE STATUS','MDG_ID':'MDGLC'}
f_gens.rename(columns = g, inplace = True)
f_sites.rename(columns = s, inplace = True)
f_cells = pd.read_excel("/Users/txvance/Documents/PSPS/OpsTracker_Raw_Files/NorCal_CellInfo.xlsx", usecols=['PSLC', 'eNodeB'])
f_cells5g = pd.read_excel("/Users/txvance/Documents/PSPS/OpsTracker_Raw_Files/norcal_cell_info_5g.xlsx", usecols=['PSLC', 'GNODEB'])

# Static files to capture sites that PGE provided but aren't in OT yet. Will result in some duplication, including sites with multiple meters
f_vzb = pd.read_excel("/Users/txvance/Documents/PSPS/OpsTracker_Raw_Files/PSPS_VZB_Sites.xlsx")
f_engie = pd.read_excel("/Users/txvance/Documents/PSPS/OpsTracker_Raw_Files/PSPS_Engie_unmatched_Sites.xlsx")

# merging the files using PSLC as the index. There are some duplicates in gen and sites files, lots of duplicates in the cell info because of B2B and 5G gNodeBs
# documentation on pandas merge https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.merge.html?highlight=merge#pandas.DataFrame.merge
f_merged = f_sites.merge(f_gens, left_on="PSLC", right_on="PSLC", how="left")
f_merged = f_merged.merge(f_cells, left_on="PSLC", right_on="PSLC", how="left")
f_merged = f_merged.merge(f_cells5g, left_on="PSLC", right_on="PSLC", how="left")

# same function but only combining the sites, gens and PSPS/PGE files for Ops
f_merged_ops = f_sites.merge(f_gens, left_on="PSLC", right_on="PSLC", how="left")
# create a new dataframe to concatenate the merged data with a static VZB/VZS file of meter numbers from PGE for non-wireless locations
# requires that the PSPS_VZB_Sites.xlsx file exist in the directory, same format as PSPS_MAIN, but with the random meters PGE provides for non-VZW locations in scope
# uses pd.concat https://pandas.pydata.org/pandas-docs/stable/user_guide/merging.html
# Remove the Engie and PGE unmatched after Gennie cleaned up OT 04/28/2022
frames = [f_merged_ops, f_vzb, f_engie]
concat_ops = pd.concat(frames)
concat_ops['GEN Y/N'] = concat_ops['GEN Y/N'].fillna(0)
concat_ops['PLUG Y/N'] = concat_ops['PLUG Y/N'].fillna(0)
concat_ops['RM Y/N'] = concat_ops['RM Y/N'].fillna(0)
concat_ops['M/W HUB Y/N'] = concat_ops['M/W HUB Y/N'].fillna(0)
concat_ops['HUB Y/N'] = concat_ops['HUB Y/N'].fillna(0)

# https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.replace.html

# Bulk replace 0 for No and 1 for yes in the various columns
map_dict = {0:'NO', 1:'YES',"VZB FACILITY":"VZB FACILITY", "FRONTIER FACILITY":"FRONTIER FACILITY","VZW RETAIL SALES":"VZW RETAIL SALES", "Operational":"YES","Y":"YES","N":"NO","NO":"NO","YES":"YES", "Diesel":"Diesel", "Propane":"Propane", "Pacific Gas & Electric":"PG&E", "PGE":"PG&E", "SCE":"SCE", "Southern California Edison":"SCE", "So Cal Edison":"SCE","SO CAL EDISON":"SCE","Southern Cal. Edison":"SCE" }
#concat_ops['PG&E Fee Property'] = concat_ops['PG&E Fee Property'].map(map_dict)
concat_ops['FUEL TYPE'] = concat_ops['FUEL TYPE'].map(map_dict)
concat_ops['POWER COMPANY'] = concat_ops['POWER COMPANY'].map(map_dict)
concat_ops['RM Y/N'] = concat_ops['RM Y/N'].map(map_dict)
concat_ops['M/W HUB Y/N'] = concat_ops['M/W HUB Y/N'].map(map_dict)
concat_ops['HUB Y/N'] = concat_ops['HUB Y/N'].map(map_dict)

map_dict_gen = {0:"NO", 1:"YES", "Operational":"YES", "Non-operational":"NO","Not Operational":"NO", "Y":"YES","N":"NO","NO":"NO","YES":"YES","VZB FACILITY":"VZB FACILITY","FRONTIER FACILITY":"FRONTIER FACILITY", "VZW RETAIL SALES":"VZW RETAIL SALES"}
concat_ops['GEN Y/N'] = concat_ops['GEN Y/N'].map(map_dict_gen)
concat_ops['PLUG Y/N'] = concat_ops['PLUG Y/N'].map(map_dict_gen)

# Format the PSPS_MAIN_SCE sheet for Ops
# establish the xlsxwriter functionality, defining "writer" as the variable for the workbook filename
writer = pd.ExcelWriter('/Users/txvance/Documents/PSPS/Tracker/PSPS_MAIN_SCE.xlsx', engine='xlsxwriter')
# Create the merged sheet and output to the file name based on the writer variable
concat_ops.to_excel(writer, index=False, sheet_name='PSPS_MAIN_SCE',columns=['POWER METER','NOTES', 'POWER COMPANY','FIRE TIER', 'PSPS PROB', 'MDGLC','PSLC', 'SITE NAME', 'ADDRESS','CITY','COUNTY', 'GEN Y/N', 'PLUG Y/N', 'PLUG TYPE', 'GEN SIZE', 'FUEL TYPE', 'TANK SIZE', 'RM Y/N', 'HUB Y/N','M/W HUB Y/N', 'FIELD ENGINEER','OPS MANAGER', 'SITE STATUS', 'NOTES'])
# Establish the workbook variable
workbook = writer.book

# Setup some formating definitions
# formatting for any cells/columns that need to be center justified
header_format = workbook.add_format()
header_format.set_bold()
header_format.set_align('center')
header_format.set_text_wrap()

cell_format_center = workbook.add_format()
cell_format_center.set_align('center')
cell_format_center.set_text_wrap()
cell_format_left = workbook.add_format()
cell_format_left.set_align('left')
cell_format_left.set_text_wrap()
# Define the worksheet variable
worksheet = writer.sheets['PSPS_MAIN_SCE']
# Apply some formatting to groups of columns, including cell width and applying the cell formatting previously defined as appropriate
worksheet.set_row(0, None, header_format)
worksheet.set_column('A:A', 16, cell_format_center)
worksheet.set_column('B:B', 12, cell_format_center)
worksheet.set_column('C:C', 20, cell_format_center)
worksheet.set_column('D:D', 9, cell_format_center)
worksheet.set_column('E:E', 8, cell_format_center)
worksheet.set_column('F:F', 14, cell_format_center)
worksheet.set_column('G:G', 9, cell_format_center)
worksheet.set_column('H:H', 37, cell_format_left)
worksheet.set_column('I:I', 44, cell_format_left)
worksheet.set_column('J:J', 22, cell_format_left)
worksheet.set_column('K:K', 16.5, cell_format_left)
worksheet.set_column('L:M', 11, cell_format_center)
worksheet.set_column('N:N', 9,  cell_format_center)
worksheet.set_column('O:P', 14, cell_format_center)
worksheet.set_column('Q:S', 9,  cell_format_center)
worksheet.set_column('T:T', 20, cell_format_center)
worksheet.set_column('U:U', 16, cell_format_left)
worksheet.set_column('V:V', 17, cell_format_center)
worksheet.set_column('W:AB', 19, cell_format_center)
# Set some worksheet formatting, including creating filter dropdowns and freeze the top row
worksheet.freeze_panes(1, 0)
worksheet.autofilter('A1:AC9999')
# Save the sheet, using new command for panda as writer.save() was deprecated 2023/06
writer.close()

