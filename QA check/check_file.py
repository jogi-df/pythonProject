import pandas as pd
import easygui
import openpyxl as op
import numpy as np
import os
#from functions import check_optional_col, check_required_col, opt_in, opt_in_check, check_phone_col, whitespace_remover, check_exists
#from validate_email import validate_email
from io import StringIO
import requests

df1 = pd. DataFrame()

def follow_url(url_wrapped):
    response = requests.get(url_wrapped)
    return response.url

working_folder = easygui.diropenbox(msg="Select the working folder", title="Working folder")
fold = StringIO(working_folder)


#easygui requests a file select for the main data frame.  Easygui returns the path of the file
raw_file_path = easygui.fileopenbox(msg="Select the raw data file.  Don't use the original!", default=working_folder + "\*")
file = StringIO(raw_file_path)
df = pd.read_excel(raw_file_path)

df = df.replace(r"<|>", r'"', regex=True)
df = df.replace(r"^ ", r"", regex=True)

print(df)

#df.loc[df['Email'].str.match('^"http'), 'Email'] = follow_url(df.loc[df['Email'].str.match('^http')])
#df1['earl'] = df.loc[df['Email'].str.match('^http')]

#print(df1)

#for i in range(len(df)):
#    if df[df['Email'].str.match('^http')]:
#        df['Email'] = follow_url(df['Email'])


#print(df)

df.to_excel(r"C:\Users\jogi\OneDrive - binary-tech.com\Consulting\DF 2022\QA check\working_copy_urls.xlsx", index=False)