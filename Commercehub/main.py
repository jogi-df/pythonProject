import pandas as pd
import easygui
import openpyxl as op
import numpy as np
import os
from io import StringIO

df1 = pd. DataFrame()

conf = pd.read_excel(r"C:\Users\jogi\OneDrive - binary-tech.com\Consulting\Commercehub\connections_on_both_platforms.xlsx", 'connections_on_both_platforms' )


xls=pd.ExcelFile(r"C:\Users\jogi\OneDrive - binary-tech.com\Consulting\Commercehub\BI Data (CommerceHub & DSCO Active) 10-24-2022.xlsx")


#sets the working folder
#working_folder = easygui.diropenbox(msg="Select the working folder", title="Working folder")
#fold = StringIO(working_folder)


#easygui requests a file select for the main data frame.  Easygui returns the path of the file
#raw_file_path = easygui.fileopenbox(msg="Select the raw data file.  Don't use the original!", default=working_folder + "\*")
#file = StringIO(raw_file_path)
#xls=pd.ExcelFile(raw_file_path)


df=pd.read_excel(xls, 'DSCO Suppliers')
df2=pd.read_excel(xls, 'Before & After by Supplier')


df1[['Supplier ID', 'Company Name', 'Supplier Account Name']] = df[['Supplier ID', 'Company Name', 'Supplier Name']]

df1['Domain'] = df['Contact Email'].str.split('@').str[1]

df1[['Address1', 'City', 'State', 'Country', 'Status']] = df[['Supplier Address', 'Supplier City', 'Supplier State', 'Supplier Country', 'Supplier Status']]

df1[['Retailer Name', 'Supplier Trading Partner ID', 'Supplier Trading Partner Name', 'DSCO Last 12 Month Order']] = df[['Retailer Name', 'Supplier Trading Partner ID','Supplier Trading Partner Name', 'Last 12 Month Order Total']]
#print(df1.columns)

#match to connection on both platforms file--supplier id is key
df1['ODD-ID'] = df1['Supplier ID'].map(conf.drop_duplicates().set_index('supplier_id')['supplierOrgId'])
df1['supplierOrgName'] = df1['Supplier ID'].map(conf.drop_duplicates().set_index('supplier_id')['supplierOrgName'])
df1['supplier_actual_name'] = df1['Supplier ID'].map(conf.drop_duplicates().set_index('supplier_id')['supplier_actual_name'])
df1['retailer_name'] = df1['Supplier ID'].map(conf.drop_duplicates().set_index('supplier_id')['retailer_name'])
df1['retailer_id'] = df1['Supplier ID'].map(conf.drop_duplicates().set_index('supplier_id')['retailer_id'])

#join to before and after by supplier ws
#df1['ODD-ID'] = df1['Supplier Name'].map(df2.drop_duplicates().set_index('Supplier Org')['Supplier ODD ID'])

#save df file to temp excel doc for openpyxl
#df1.to_excel(working_folder + "\working_copy.xlsx", index=False)
df1.to_excel(r"C:\Users\jogi\OneDrive - binary-tech.com\Consulting\Commercehub\working_copy.xlsx", index=False)