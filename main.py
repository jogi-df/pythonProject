import pandas
import easygui


#easygui requests a file select for the main data frame.  Easygui returns the path of the file
#from io import StringIO
#raw_file_path = easygui.fileopenbox(msg="Select the raw data file.  Don't use the original!")
#file = StringIO(raw_file_path)
#df = pandas.read_excel(raw_file_path)
df = pandas.read_excel(r"C:\test\testrawdata.xlsx")

# these are from the Adobe Lead Import Template Worldwide tabs and are used to compare against
#from io import StringIO
#raw_file_path1 = easygui.fileopenbox(msg="Select the template data file.  Don't use the original!")
#file = StringIO(raw_file_path1)
#df_acceptable = pandas.read_excel(raw_file_path1, sheet_name="Acceptable List of Values")
#df_states = pandas.read_excel(raw_file_path1, sheet_name="Territory State List")
df_acceptable = pandas.read_excel(r'C:\test\Lead Import Template Worldwide.xlsx', sheet_name="Acceptable List of Values")
df_states = pandas.read_excel(r'C:\test\Lead Import Template Worldwide.xlsx', sheet_name="Territory State List")

#check CID length for 18 characters
df.loc[df['CID'].apply(len) == 18, 'CID_status'] = 'TRUE'
df.loc[df['CID'].apply(len) != 18, 'CID_status'] = 'FALSE'


#permissions date cleanup
df["Permissions Create Date"] = pandas.to_datetime(df["Permissions Create Date"]).dt.strftime("%m%d%Y")

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

#check if state is 2 characters - probably don't need because it compares against the template list
#df['State_length'] = df['State'].apply(len)
#df.loc[df['State_length'] == 2, 'State_status'] = 'State Good'
#df.loc[df['State_length'] != 2, 'State_status'] = 'Check State'

#check if state abbreviation is in list (US only)

boolvalue = df['State'].isin(df_states['Abbreviation'])
df['State_Abbr_Good'] = boolvalue

#change OPT IN Email to Y N U
df.loc[df['OPT IN EMAIL'] == 'Yes', 'OPT IN EMAIL'] = 'Y'
df.loc[df['OPT IN EMAIL'] == 'No', 'OPT IN EMAIL'] = 'N'
df.loc[df['OPT IN EMAIL'].isnull(), 'OPT IN EMAIL'] = 'U'

#change OPT IN mail to Y N U
df.loc[df['OPT IN MAIL'] == 'Yes', 'OPT IN MAIL'] = 'Y'
df.loc[df['OPT IN MAIL'] == 'No', 'OPT IN MAIL'] = 'N'
df.loc[df['OPT IN MAIL'].isnull(), 'OPT IN MAIL'] = 'U'

#change OPT IN phone to Y N U
df.loc[df['OPT IN PHONE'] == 'Yes', 'OPT IN PHONE'] = 'Y'
df.loc[df['OPT IN PHONE'] == 'No', 'OPT IN PHONE'] = 'N'
df.loc[df['OPT IN PHONE'].isnull(), 'OPT IN PHONE'] = 'U'

#change OPT IN THIRD PARTY to Y N U
df.loc[df['OPT IN THIRD PARTY'] == 'Yes', 'OPT IN THIRD PARTY'] = 'Y'
df.loc[df['OPT IN THIRD PARTY'] == 'No', 'OPT IN THIRD PARTY'] = 'N'
df.loc[df['OPT IN THIRD PARTY'].isnull(), 'OPT IN THIRD PARTY'] = 'U'

#check if product interest is spelled correctly
boolvalue = df['Product Interest'].isin(df_acceptable['Product Interest'])
df['Product_Good'] = boolvalue

#check if estimated number of units is formatted correctly
#df.loc[df['Estimated Number of Units'].isnull(), 'Est # Units'] = ''
#boolvalue = df['Estimated Number of Units'].isin(df_acceptable['Estimated Number of Units'])
#df['Est # Units'] = boolvalue

#df.loc[df['Estimated Number of Units'].notnull(), df['Est Units']] = df['Estimated Number of Units'].isin(df_acceptable['Estimated Number of Units'])
#boolvalue1 = df['Estimated Number of Units'].isin(df_acceptable['Estimated Number of Units'])
boolvalue1 = df['Estimated Number of Units'].isnull()
#df['Est # Units1'].all = boolvalue1


#if df['Est # Units1'].all() == "True":
#   df['Est # Units2'] = "blank"
#df.loc[df['Est # Units1'] == 'TRUE', 'Est # Units1'] = 'blank'

#check if industry is in list
boolvalue = df['Industry'].isin(df_acceptable['Industry'])
df['Industry_Good'] = boolvalue

#check if functional area is in list
boolvalue = df['Functional Area/Department'].isin(df_acceptable['Functional Area/Department'])
df['Industry_Good'] = boolvalue

#Job Function
boolvalue = df['Job Function'].isin(df_acceptable['Job Function'])
df['Job_Function_Good'] = boolvalue

# check employee range and fix if date is showing and then compare to acceptable values
df.loc[df['Employee Range'] == 'Oct-99', 'Employee Range'] = '10-99'
df.loc[df['Employee Range'] == '9-Jan', 'Employee Range'] = '1-9'

boolvalue = df['Employee Range'].isin(df_acceptable['Employee Range'])
df['Employee_Range_Good'] = boolvalue

#check if required fields exist
#placeholder

#save df file to temp excel doc for openpyxl
df.to_excel(r'C:\test\testrawdata_filtered_temp.xlsx', index=False)

#find any columns with 'Field' in the name and create a df and drop if empty
df_field = pandas.DataFrame()
df_field = df.filter(like='Field', axis=1)
df_field.dropna(how='all', axis=1, inplace=True)
pandas.concat([df, df_field], axis='columns')

#openpyxl conditional formatting section
from openpyxl import load_workbook
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting import Rule
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule

wb = load_workbook(r'C:\test\testrawdata_filtered_temp.xlsx')
ws = wb.active

for r in dataframe_to_rows(df, index=True, header=True):
    ws.append(r)

red_fill = PatternFill(bgColor="FFC7CE")
dxf = DifferentialStyle(fill=red_fill)
rule = Rule(type="containsText", operator="containsText", text="highlight", dxf=dxf)
rule.formula = ['NOT(ISERROR(SEARCH("FALSE",AV2)))']
ws.conditional_formatting.add('AV2:BD400', rule)

wb.save(r'C:\test\testrawdata_filtered.xlsx')

#now lets copy the columns from raw to the template but I only know how to do it in pandas
df_template = pandas.read_excel(r'c:\test\Lead Import Template Worldwide.xlsx', sheet_name="Template (File Format)")
df_filtered = pandas.read_excel(r'c:\test\testrawdata_filtered.xlsx')
df_new = pandas.DataFrame()


df_new = df_template.copy()
df_new['Campaign ID'] = df_filtered['CID'].copy()
df_new['Permissions Create Date '] = df_filtered['Permissions Create Date'].copy()
df_new['Email'] = df_filtered['Email'].copy()
df_new['Zip'] = df_filtered['Zip'].copy()
df_new['Country'] = df_filtered['Country'].copy()
df_new['Phone'] = df_filtered['Phone'].copy()
df_new['First Name'] = df_filtered['First Name'].copy()
df_new['Last Name'] = df_filtered['Last Name'].copy()
df_new['Company Name'] = df_filtered['Company Name'].copy()
df_new['OPT IN EMAIL'] = df_filtered['OPT IN EMAIL'].copy()
df_new['OPT IN MAIL'] = df_filtered['OPT IN MAIL'].copy()
df_new['OPT IN PHONE'] = df_filtered['OPT IN PHONE'].copy()
df_new['Product Interest'] = df_filtered['Product Interest'].copy()
df_new['Industry'] = df_filtered['Industry'].copy()
df_new = pandas.concat([df_new, df_field], axis=1)

#print (df_template)
df_new.to_excel(r'C:\test\testrawdata_final.xlsx', index=False)


#save file to excel format
#df.to_excel(r'C:\test\testrawdata_filtered.xlsx', index=False)
#print(df['Zip_status'])
