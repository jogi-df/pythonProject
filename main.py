import pandas as pd
import easygui
import openpyxl as op
import os
from functions import check_optional_col, check_required_col, opt_in, check_exists
from validate_email import validate_email
from io import StringIO

#reference files
df_zip = pd.read_excel(r'C:\test\ZIP_Locale_Detail.xls')


#easygui requests a file select for the main data frame.  Easygui returns the path of the file
#from io import StringIO
#working_folder_path = easygui.fileopenbox(msg="Select the raw data file.  Don't use the original!")
#file = StringIO(working_folder_path)


#easygui requests a file select for the main data frame.  Easygui returns the path of the file
raw_file_path = easygui.fileopenbox(msg="Select the raw data file.  Don't use the original!")
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


#check CID length for 18 characters
df.loc[df['Campaign ID'].apply(len) == 18, 'CID_status'] = 'TRUE'
df.loc[df['Campaign ID'].apply(len) != 18, 'CID_status'] = 'FALSE'

#permissions date cleanup
df.loc[df['Permissions Create Date '].isnull(), 'Perm Date'] = 'required'

#print(df['Permissions Create Date'])
#if df["Permissions Create Date"].all().str.contains('/'):
df["Permissions Create Date "] = pd.to_datetime(df["Permissions Create Date "]).dt.strftime("%m%d%Y")


check_optional_col('Attended','Attend',df,df_acceptable)

#check if email address is in a standard format
df['Email_Good'] = df['Email'].apply(validate_email)

check_optional_col('Salutation','Sal Good',df,df_acceptable)

check_exists('First Name', 'F Name',df)

check_exists('Last Name', 'L Name', df)

check_exists('Company Name', 'Comp Good', df)

check_required_col('State','State Good',df,df_states,'Abbreviation')

#check zip code length for 5 characters - pad to 5 with leading zero
#df['Zip'] = df['Zip'].astype(str)
#df['Zip'] = df['Zip'].str.zfill(5)

#check_required_col('Zip','Zip Good',df,df_zip,'DELIVERY ZIPCODE')
df[df['Zip'].str.match("^[0-9]{5}(?:-[0-9]{4})?$") == True]


#If Regex.IsMatch(df['Zip'], "^[0-9]{5}(?:-[0-9]{4})?$") Then
#    Console.WriteLine("Valid ZIP code")
#Else
#    Console.WriteLine("Invalid ZIP code")
#End If

#zip_length = df['Zip'].astype(str)
#df['Zip_Length'] = zip_length.str.len()
#df.loc[df['Zip_Length'] == 5, 'Zip_status'] = 'TRUE'
#df.loc[df['Zip_Length'] != 5, 'Zip_status'] = 'FALSE'

#if country is United States, replace with US - required
check_required_col('Country','Country Good',df,df_acceptable,'Country')
#df.loc[df['Country'] == "United States", 'Country'] = 'US'

#clean up the opt in responses - required
opt_in(df,'OPT IN EMAIL')
opt_in(df,'OPT IN MAIL')
opt_in(df,'OPT IN PHONE')
opt_in(df,'OPT IN THIRD PARTY')
#df.loc[df['Country'] == "United States", 'Country'] = 'US'

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
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule

wb = load_workbook(r'C:\test\working_copy.xlsx')
ws = wb.active

#for r in dataframe_to_rows(df, index=True, header=True):  #why is this in here???
#    ws.append(r)

red_fill = PatternFill(bgColor="FFC7CE")
dxf = DifferentialStyle(fill=red_fill)
rule = Rule(type="containsText", operator="containsText", text="highlight", dxf=dxf)
rule.formula = ['NOT(ISERROR(SEARCH("FALSE",AV2)))']
rule1 = Rule(type="containsText", operator="containsText", text="highlight", dxf=dxf)
rule1.formula = ['NOT(ISERROR(SEARCH("required",AV2)))']
ws.conditional_formatting.add('AV2:BT4000', rule)
ws.conditional_formatting.add('AV2:BT4000', rule1)
wb.save(r'C:\test\working_copy.xlsx')



#open template workbook, delete all tabs except template sheet and save as new file
from openpyxl import load_workbook, Workbook

wb = load_workbook(r"C:\test\Lead Import Template Worldwide.xlsx")

sheets = wb.sheetnames

for s in sheets:
    if s != 'Template (File Format)':
        sheet_name = wb[s]
        wb.remove(sheet_name)

wb.save(r'c:\test\new.xlsx')



#now lets copy the columns from raw to the template

wbt = op.load_workbook(r'c:\test\new.xlsx')
wst = wbt['Template (File Format)']

wbf = op.load_workbook(r'c:\test\working_copy.xlsx')
wsf = wbf['Sheet1']

wb3 = op.load_workbook(r'c:\test\col_map.xlsx')
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

wbt.save(r'c:\test\final.xlsx')

#clean up temporary template file
if os.path.exists(r"c:\test\new.xlsx"):
    os.remove(r"c:\test\new.xlsx")


