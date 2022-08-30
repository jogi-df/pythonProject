import pandas as pd
import easygui
import openpyxl as op
import os

def check_optional_col(colchk,colres,datasrc,datatemplate):          #this checks the required columns filtering out the blanks

    #df = pd.read_excel(r"C:\test\testrawdata.xlsx")
    #dfa = pd.read_excel(r'C:\test\Lead Import Template Worldwide.xlsx', sheet_name="Acceptable List of Values")

    dftemp = datasrc[[colchk]].copy()

    dftemp1 = dftemp.copy()
    dftemp1.dropna(inplace = True)

    dftempa = datatemplate[[colchk]].copy()
    dftempa.dropna(inplace = True)
    dftempa = dftempa.copy()

    dftemp[colres] = dftemp1[colchk].isin(dftempa[colchk])
    dftemp[colres].fillna("optional", inplace=True)

    datasrc[colres] = dftemp[colres]

    #df.to_excel(r'C:\test\working_copy.xlsx', index=False)

def check_required_col(colchk,colres,datasrc,datatemplate):          #this checks the required columns filtering out the blanks

    #df = pd.read_excel(r"C:\test\testrawdata.xlsx")
    #dfa = pd.read_excel(r'C:\test\Lead Import Template Worldwide.xlsx', sheet_name="Acceptable List of Values")

    dftemp = datasrc[[colchk]].copy()

    dftemp1 = dftemp.copy()
    dftemp1.dropna(inplace = True)

    dftempa = datatemplate[[colchk]].copy()
    dftempa.dropna(inplace = True)
    dftempa = dftempa.copy()

    dftemp[colres] = dftemp1[colchk].isin(dftempa[colchk])
    dftemp[colres].fillna("required", inplace=True)

    datasrc[colres] = dftemp[colres]

    #df.to_excel(r'C:\test\working_copy.xlsx', index=False)


def opt_in(df,name):                            #cleans up optin values
    df.loc[df[name] == 'Yes', name] = 'Y'
    df.loc[df[name] == 'No', name] = 'N'
    df.loc[df[name].isnull(), name] = 'U'
    return df

#easygui requests a file select for the main data frame.  Easygui returns the path of the file
#from io import StringIO
#working_folder_path = easygui.fileopenbox(msg="Select the raw data file.  Don't use the original!")
#file = StringIO(working_folder_path)


#easygui requests a file select for the main data frame.  Easygui returns the path of the file
from io import StringIO
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
df.loc[df['CID'].apply(len) == 18, 'CID_status'] = 'TRUE'
df.loc[df['CID'].apply(len) != 18, 'CID_status'] = 'FALSE'


#permissions date cleanup
df["Permissions Create Date"] = pd.to_datetime(df["Permissions Create Date"]).dt.strftime("%m%d%Y")

#Salutation format check
check_optional_col('Salutation','Sal Good',df,df_acceptable)

#first name check if empty
#boolvalue = df['First Name'].notnull
#df['Salutation_Good'] = boolvalue


#check if email address is in a standard format
from validate_email import validate_email
email_valid = df['Email'].apply(validate_email)
df['Email_Good'] = email_valid


#check zip code length for 5 characters - pad to 5 with leading zero
df['Zip'] = df['Zip'].astype(str)
df['Zip'] = df['Zip'].str.zfill(5)

zip_length = df['Zip'].astype(str)
df['Zip_length'] = zip_length.str.len()
df.loc[df['Zip_length'] == 5, 'Zip_status'] = 'TRUE'
df.loc[df['Zip_length'] != 5, 'Zip_status'] = 'FALSE'


#if country is United States, replace with US

df.loc[df['Country'] == "United States", 'Country'] = 'US'


#check if state abbreviation is in list (US only)
#check_required_col('State','Prod Int',df,df_states)

boolvalue = df['State'].isin(df_states['Abbreviation'])
df['State_Abbr_Good'] = boolvalue

#clean up the opt in responses
opt_in(df,'OPT IN EMAIL')
opt_in(df,'OPT IN MAIL')
opt_in(df,'OPT IN PHONE')
opt_in(df,'OPT IN THIRD PARTY')



#check if product interest is spelled correctly
check_required_col('Product Interest','Prod Int',df,df_acceptable)

#check if estimated number of units is formatted correctly
check_optional_col('Estimated Number of Units','Est Num Units',df,df_acceptable)



#check if industry is in list - required
check_required_col('Industry','Industry_good',df,df_acceptable)
#boolvalue = df['Industry'].isin(df_acceptable['Industry'])
#df['Industry_Good'] = boolvalue

#check if functional area is in list - optional
check_optional_col('Functional Area/Department','Funct Area',df,df_acceptable)
#boolvalue = df['Functional Area/Department'].isin(df_acceptable['Functional Area/Department'])
#df['Industry_Good'] = boolvalue

#Job Function - optional
check_optional_col('Job Function','Job Funct',df,df_acceptable)
#boolvalue = df['Job Function'].isin(df_acceptable['Job Function'])
#df['Job_Function_Good'] = boolvalue

# check employee range and fix if date is showing and then compare to acceptable values
df.loc[df['Employee Range'] == 'Oct-99', 'Employee Range'] = '10-99'
df.loc[df['Employee Range'] == '9-Jan', 'Employee Range'] = '1-9'

check_required_col('Employee Range','emp range good',df,df_acceptable)
#boolvalue = df['Employee Range'].isin(df_acceptable['Employee Range'])
#df['Employee_Range_Good'] = boolvalue


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
ws.conditional_formatting.add('AV2:BD400', rule)

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


