import pandas as pd
import easygui
import openpyxl as op
import os
from functions import check_optional_col, check_required_col, opt_in, check_exists, check_phone_col
from validate_email import validate_email
from io import StringIO

#reference files
df_zip = pd.read_excel(r'C:\test\ZIP_Locale_Detail.xls')



#sets the working folder
working_folder = easygui.diropenbox(msg="Select the working folder", title="Working folder")
fold = StringIO(working_folder)


#easygui requests a file select for the main data frame.  Easygui returns the path of the file
raw_file_path = easygui.fileopenbox(msg="Select the raw data file.  Don't use the original!", default=working_folder + "\*")
file = StringIO(raw_file_path)
df = pd.read_excel(raw_file_path)
#df = pd.read_excel(r"C:\test\testrawdata.xlsx")   #uncomment this for testing and comment out everything else

# these are from the Adobe Lead Import Template Worldwide tabs and are used to compare against
#from io import StringIO
#raw_file_path1 = easygui.fileopenbox(msg="Select the template data file.  Don't use the original!")
#file = StringIO(raw_file_path1)
#df_acceptable = pd.read_excel(raw_file_path1, sheet_name="Acceptable List of Values")
#df_states = pd.read_excel(raw_file_path1, sheet_name="Territory State List")
df_acceptable = pd.read_excel(r'C:\test\Lead Import Template Worldwide.xlsx', sheet_name="Acceptable List of Values")          #uncomment this for testing and comment out everything else
df_states = pd.read_excel(r'C:\test\Lead Import Template Worldwide.xlsx', sheet_name="Territory State List")                   #uncomment this for testing and comment out everything else

#remove any trailing whitespaces in the column names (shows up in permissions create date)
df.columns = df.columns.str.rstrip()


#sometimes raw data comes in on the template where Campaign ID is the CID column.  Lets normalize the name.
if 'Campaign ID' in df.columns:
    df.rename(columns = {"Campaign ID":"CID"}, inplace = True)


#check CID length for 18 characters
df.loc[df['CID'].apply(len) == 18, 'CID_status'] = 'TRUE'
df.loc[df['CID'].apply(len) != 18, 'CID_status'] = 'FALSE'

#permissions date cleanup
#df.loc[df['Permissions Create Date'].isnull(), 'Perm Date'] = 'required'

#print(df['Permissions Create Date'])
#if df["Permissions Create Date"].all().str.contains('/'):
#df["Permissions Create Date"] = pd.to_datetime(df["Permissions Create Date"]).dt.strftime("%m%d%Y")
df['Permissions Create Date'] = df['Permissions Create Date'].astype(str)
df['Perm Date'] = df['Permissions Create Date'].str.match("[0-9]{8}")




check_optional_col('Attended','Attend',df,df_acceptable)

#check if email address is in a standard format
df['Email_Good'] = df['Email'].apply(validate_email)

check_optional_col('Salutation','Sal Good',df,df_acceptable)

check_exists('First Name', 'F Name',df)

check_exists('Last Name', 'L Name', df)

check_exists('Company Name', 'Comp Good', df)

check_required_col('State','State Good',df,df_states,'Abbreviation')

#check zip code length for 5 characters or 5+4
df['Zip'] = df['Zip'].astype(str)
df['Zip Status'] = df['Zip'].str.match("^[0-9]{5}(?:-[0-9]{4})?$")

check_required_col('Country','Country Good',df,df_acceptable,'Country')

check_phone_col('Phone', 'Ph Status', df)

#clean up the opt in responses - required
opt_in(df,'OPT IN EMAIL')
opt_in(df,'OPT IN MAIL')
opt_in(df,'OPT IN PHONE')
opt_in(df,'OPT IN THIRD PARTY')

check_required_col('Product Interest','Prod Int',df,df_acceptable, 'Product Interest')

check_optional_col('Additional Product Interest','Addl Prod Int',df,df_acceptable)

check_optional_col('Estimated Number of Units','Est Num Units',df,df_acceptable)

check_optional_col('Timeline for Purchasing','Timeline',df,df_acceptable)

check_required_col('Industry','Industry_good',df,df_acceptable, 'Industry')

check_optional_col('Functional Area/Department','Funct Area',df,df_acceptable)

check_optional_col('Job Function','Job Funct',df,df_acceptable)

# check employee range and fix if date is showing and then compare to acceptable values
#df.loc[df['Employee Range'] == 'Oct-99', 'Employee Range'] = '10-99'
#df.loc[df['Employee Range'] == '9-Jan', 'Employee Range'] = '1-9'
check_required_col('Employee Range','emp range good',df,df_acceptable, 'Employee Range')

check_optional_col('Revenue Range','Revenue Rg',df,df_acceptable)

check_optional_col('Budget Established','Bud Est',df,df_acceptable)

check_optional_col('Request Follow-Up/Demo','Req FU',df,df_acceptable)

#save df file to temp excel doc for openpyxl
df.to_excel(r'C:\test\working_copy.xlsx', index=False)


#make a copy of df and find any columns with 'Field' in the name and create a df and drop if empty
#how to use this??
df_field = df.copy()
df_field = df_field.filter(like='Field', axis=1)
df_field.dropna(how='all', axis=1, inplace=True)
#df_field1 =pd.concat([df, df_field], axis='columns')
#print(df_field)


#openpyxl conditional formatting section
from openpyxl import load_workbook
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting import Rule

wb = load_workbook(working_folder + "\working_copy.xlsx")
ws = wb.active

red_fill = PatternFill(bgColor="FFC7CE")
dxf = DifferentialStyle(fill=red_fill)
rule = Rule(type="containsText", operator="containsText", text="highlight", dxf=dxf)
rule.formula = ['NOT(ISERROR(SEARCH("FALSE",AP2)))']
rule1 = Rule(type="containsText", operator="containsText", text="highlight", dxf=dxf)
rule1.formula = ['NOT(ISERROR(SEARCH("required",AP2)))']
rule2 = Rule(type="containsText", operator="containsText", text="highlight", dxf=dxf)
rule2.formula = ['NOT(ISERROR(SEARCH("set OIP N",AP2)))']
ws.conditional_formatting.add('AP2:BT4000', rule)
ws.conditional_formatting.add('AP2:BT4000', rule1)
ws.conditional_formatting.add('AP2:BT4000', rule2)
wb.save(r'C:\test\working_copy.xlsx')



#open template workbook, delete all tabs except template sheet and save as new file
from openpyxl import load_workbook, Workbook

wb = load_workbook(working_folder + "\Lead Import Template Worldwide.xlsx")

sheets = wb.sheetnames

for s in sheets:
    if s != 'Template (File Format)':
        sheet_name = wb[s]
        wb.remove(sheet_name)

wb.save(working_folder + "\\new.xlsx")



#now lets copy the columns from raw to the template

wbt = op.load_workbook(working_folder + "\\new.xlsx")
wst = wbt['Template (File Format)']

wbf = op.load_workbook(working_folder + "\\working_copy.xlsx")
wsf = wbf['Sheet1']

wb3 = op.load_workbook(working_folder + "\\col_map.xlsx")
ws3 = wb3['Sheet1']


# get map column ids
mr = ws3.max_row
mc = ws3.max_column

mr1 = wsf.max_row

for x in range (2, mr + 1):
    a = ws3.cell(row = x, column = 1).value
    b = ws3.cell(row = x, column = 2).value
    if a is None:
        continue
    for i in range (1, mr1 + 1):
        c = wsf.cell(row = i, column = a)
        wst.cell(row = i, column = b).value = c.value

wbt.save(working_folder + "\\final.xlsx")

#clean up temporary template file
if os.path.exists(working_folder + "\\new.xlsx"):
    os.remove(working_folder + "\\new.xlsx")


