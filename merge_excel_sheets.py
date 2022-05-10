# # 2022/02 Terry Vance (vanceter)
# OpsTracker file merge for PSPS events.
# importing the module
import pandas as pd
import xlsxwriter

file_path = "/Users/txvance/Documents/PSPS/"
file_path_raw = "/Users/txvance/Documents/PSPS/OpsTracker_Raw_Files/"

# NOTE each of the underlying xls files needs to have the PSLC value - in some of them, the column header needs to be renamed from PS Loc
# Also need to make sure you export OpsTracker files with the file name option checked, and change the filter to get -all- sites
# Raw data files need to be here: /Users/txvance/Documents/PSPS/OpsTracker_Raw_Files/
# Output files will go here: /Users/txvance/Documents/PSPS/Tracker/
 
# reading only the columns needed from each file
# documentation on pandas read_excel https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html
f_sites = pd.read_excel("/Users/txvance/Documents/PSPS/OpsTracker_Raw_Files/opstracker_sites.xlsx", usecols=['SITE_NAME','ADDRESS','CITY','COUNTY','PSLC','POWER_METER', 'GEN_PORTABLE_PLUG', 'GEN_PORTABLE_PLUG_TYPE', 'GO95_FIRE_ZONE_SECTOR', 'SITE_STATUS', 'IS_HUB','IS_HUB_MICROWAVE','REMOTE_MONITORING','SITETECH_NAME','SITEMGR_NAME', 'POWER_COMPANY'])
f_gens = pd.read_excel("/Users/txvance/Documents/PSPS/OpsTracker_Raw_Files/opstracker_generators.xlsx", usecols=['PSLC', 'GEN_STATUS', 'SERIALNUM', 'FUEL_TANK1', 'FUEL_TYPE1', 'MANUFACTURER', 'MODEL', 'GEN_SIZE'])
f_cells = pd.read_excel("/Users/txvance/Documents/PSPS/OpsTracker_Raw_Files/NorCal_CellInfo.xlsx", usecols=['PSLC', 'eNodeB'])
f_cells5g = pd.read_excel("/Users/txvance/Documents/PSPS/OpsTracker_Raw_Files/norcal_cell_info_5g.xlsx", usecols=['PSLC', 'GNODEB'])
f_pge = pd.read_excel("/Users/txvance/Documents/PSPS/OpsTracker_Raw_Files/PSPS_FIRE_TIER.xlsx", usecols=['PSLC', 'GO95_FIRE_ZONE_SECTOR', 'PSPS PROB', 'PG&E Fee Property'])
# Pull in static PGE master file of latest PGE list of meters that they sent at the beginning of the season, used to reconcile which meters are in scope for PSPS
f_pgemaster = pd.read_excel("/Users/txvance/Documents/PSPS/OpsTracker_Raw_Files/PGE_MASTER_LIST.xlsx", usecols=['PGE_BADGE_NUMBER', 'VLOOKUP($A1,[PSPS_MAIN.xlsx]PSPS_MAIN!A:A,1,FALSE)'])

# Static files to capture sites that PGE provided but aren't in OT yet. Will result in some duplication, including sites with multiple meters
f_vzb = pd.read_excel("/Users/txvance/Documents/PSPS/OpsTracker_Raw_Files/PSPS_VZB_Sites.xlsx")
f_unmatched = pd.read_excel("/Users/txvance/Documents/PSPS/OpsTracker_Raw_Files/PSPS_PGE_unmatched_Sites.xlsx")
f_engie = pd.read_excel("/Users/txvance/Documents/PSPS/OpsTracker_Raw_Files/PSPS_Engie_unmatched_Sites.xlsx")

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
#frames = [f_merged_ops, f_vzb, f_engie]
# Remove the Engie and PGE unmatched after Gennie cleaned up OT 04/28/2022
frames = [f_merged_ops, f_vzb, f_unmatched, f_engie]
concat_ops = pd.concat(frames)
#frames_sp = [f_merged, f_vzb, f_engie]
# Remove the Engie and PGE unmatched after Gennie cleaned up OT 04/28/2022
frames_sp = [f_merged, f_vzb, f_unmatched, f_engie]
concat_sp = pd.concat(frames_sp)
frames_pgemaster = [f_pgemaster]
concat_pgemaster = pd.concat(frames_pgemaster)

# https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.replace.html

# Bulk replace 0 for No and 1 for yes in the various columns
map_dict = {0:'NO', 1:'YES',"VZB FACILITY":"VZB FACILITY", "VZW RETAIL SALES":"VZW RETAIL SALES", "Operational":"YES","Y":"YES","N":"NO", "Diesel":"Diesel", "Propane":"Propane"}
concat_ops['PG&E Fee Property'] = concat_ops['PG&E Fee Property'].map(map_dict)
concat_ops['FUEL_TYPE1'] = concat_ops['FUEL_TYPE1'].map(map_dict)
# This command seems to null out the field, removed it 5/2/2022
#concat_ops['GEN_PORTABLE_PLUG_TYPE'] = concat_ops['GEN_PORTABLE_PLUG_TYPE'].map(map_dict)
concat_ops['REMOTE_MONITORING'] = concat_ops['REMOTE_MONITORING'].map(map_dict)
concat_ops['IS_HUB_MICROWAVE'] = concat_ops['IS_HUB_MICROWAVE'].map(map_dict)
concat_ops['IS_HUB'] = concat_ops['IS_HUB'].map(map_dict)

map_dict_gen = {0:"NO", 1:"YES", None:"NO", "Operational":"YES", "Non-operational":"NO", "Y":"YES","N":"NO"}
concat_ops['GEN_STATUS'] = concat_ops['GEN_STATUS'].map(map_dict_gen)
concat_ops['GEN_PORTABLE_PLUG'] = concat_ops['GEN_PORTABLE_PLUG'].map(map_dict_gen)
#concat_sp['PG&E Fee Property'] = concat_sp['PG&E Fee Property'].map(map_dict)
#concat_sp['GEN_STATUS'] = concat_sp['GEN_STATUS'].map(map_dict)
#concat_sp['FUEL_TYPE1'] = concat_sp['FUEL_TYPE1'].map(map_dict)
#concat_sp['GEN_PORTABLE_PLUG'] = concat_sp['GEN_PORTABLE_PLUG'].map(map_dict)
#concat_sp['GEN_PORTABLE_PLUG_TYPE'] = concat_sp['GEN_PORTABLE_PLUG_TYPE'].map(map_dict)
#concat_sp['REMOTE_MONITORING'] = concat_sp['REMOTE_MONITORING'].map(map_dict)
#concat_sp['IS_HUB_MICROWAVE'] = concat_sp['IS_HUB_MICROWAVE'].map(map_dict)
#concat_sp['IS_HUB'] = concat_sp['IS_HUB'].map(map_dict)

# creating 2 new files, the PSPS_Main for Gennie, and a version of it with eNB/gNB for SP
# Format the PSPS_MAIN sheet for Ops
# establish the xlsxwriter functionality, defining "writer" as the variable for the workbook filename
writer = pd.ExcelWriter('/Users/txvance/Documents/PSPS/Tracker/PSPS_MAIN.xlsx', engine='xlsxwriter')
# Create the merged sheet and output to the file name based on the writer variable
#f_merged_ops.to_excel(writer, index=False, sheet_name='PSPS_MAIN',columns=['POWER_METER','Fire Tier', 'PSPS PROB','PSLC', 'PG&E Fee Property', 'SITE_NAME', 'ADDRESS','CITY','COUNTY', 'GEN_STATUS','FUEL_TYPE1', 'GEN_PORTABLE_PLUG', 'GEN_PORTABLE_PLUG_TYPE', 'REMOTE_MONITORING', 'IS_HUB_MICROWAVE', 'IS_HUB','SITETECH_NAME','SITETECH_MANAGER_NAME', 'POWER_COMPANY'])
concat_ops.to_excel(writer, index=False, sheet_name='PSPS_MAIN',columns=['POWER_METER','POWER_COMPANY','GO95_FIRE_ZONE_SECTOR', 'PSPS PROB','PSLC', 'PG&E Fee Property', 'SITE_NAME', 'ADDRESS','CITY','COUNTY', 'GEN_STATUS', 'GEN_PORTABLE_PLUG', 'GEN_PORTABLE_PLUG_TYPE', 'GEN_SIZE', 'FUEL_TYPE1', 'FUEL_TANK1', 'REMOTE_MONITORING', 'IS_HUB','IS_HUB_MICROWAVE', 'SITETECH_NAME','SITEMGR_NAME', 'SITE_STATUS'])
# Add second tab to the PSPS_MAIN file to pull in the PGE master and a VLOOKUP command, used to check if anything is missing on the PSPS_MAIN, and what on the PSPS_MAIN is in scope
concat_pgemaster.to_excel(writer, index=False, sheet_name="PGEMASTERLIST", columns=['PGE_BADGE_NUMBER', 'VLOOKUP($A1,[PSPS_MAIN.xlsx]PSPS_MAIN!A:A,1,FALSE)'])
# Establish the workbook variable
workbook = writer.book

# Setup some formating definitions
# formatting for any cells/columns that need to be center justified
cell_format_center = workbook.add_format()
cell_format_center.set_align('center')


# Define the worksheet variable
worksheet = writer.sheets['PSPS_MAIN']
# Apply some formatting to groups of columns, including cell width and applying the cell formatting previously defined as appropriate
worksheet.set_column('A:B', 18, cell_format_center)
worksheet.set_column('C:E', 9, cell_format_center)
worksheet.set_column('F:F', 15, cell_format_center)
worksheet.set_column('G:G', 44)
worksheet.set_column('H:I', 22)
worksheet.set_column('J:J', 15, cell_format_center)
worksheet.set_column('K:K', 15, cell_format_center)
worksheet.set_column('L:M', 22, cell_format_center)
worksheet.set_column('N:N', 14, cell_format_center)
worksheet.set_column('O:O', 18, cell_format_center)
worksheet.set_column('P:P', 12, cell_format_center)
worksheet.set_column('Q:R', 15, cell_format_center)
worksheet.set_column('S:V', 23, cell_format_center)
worksheet.set_column('W:Y', 19, cell_format_center)
# Set some worksheet formatting, including creating filter dropdowns and freeze the top row
worksheet.freeze_panes(1, 0)
worksheet.autofilter('A1:X9999')
# Save the sheet
writer.save()

# Format the PSPS_MAIN_SP sheet for SP
# establish the xlsxwriter functionality, defining "writer" as the variable for the workbook filename
writer_sp = pd.ExcelWriter('/Users/txvance/Documents/PSPS/Tracker/PSPS_MAIN_SP.xlsx', engine='xlsxwriter')
# Create the merged sheet and output to the file name based on the writer variable
concat_sp.to_excel(writer_sp, index = False, sheet_name='PSPS_MAIN_SP', columns=['POWER_METER','PSLC', 'SITE_NAME', 'ADDRESS','CITY','COUNTY','SITETECH_NAME','SITEMGR_NAME', 'POWER_COMPANY', 'eNodeB', 'GNODEB'])

# Establish the workbook variable
workbook_sp = writer_sp.book

# Setup some formating definitions
# formatting for any cells/columns that need to be center justified
cell_format_center_sp = workbook_sp.add_format()
cell_format_center_sp.set_align('center')

# Define the worksheet variable
worksheet_sp = writer_sp.sheets['PSPS_MAIN_SP']
# Apply some formatting to groups of columns, including cell width and applying the cell formatting previously defined as appropriate
worksheet_sp.set_column('A:B', 18, cell_format_center_sp)
worksheet_sp.set_column('C:B', 10, cell_format_center_sp)
worksheet_sp.set_column('C:D', 50,)
worksheet_sp.set_column('E:E', 9)
worksheet_sp.set_column('F:F', 15)
worksheet_sp.set_column('G:H', 20)
worksheet_sp.set_column('I:I', 30)
worksheet_sp.set_column('J:K', 12, cell_format_center_sp)
# Set some worksheet formatting, including creating filter dropdowns and freeze the top row
worksheet_sp.freeze_panes(1, 0)
worksheet_sp.autofilter('A1:K9999')
# Save the sheet
writer_sp.save()