



def check_optional_col(colchk,colres,datasrc,datatemplate):
    #this checks the optional columns filtering out the blanks

    if not colchk in datasrc.columns:
        print(colchk + " column not in source")
        return

    datasrc.loc[datasrc[colchk].isnull(), colres] = 'required'

    dftemp = datasrc[[colchk]].copy()

    dftemp1 = dftemp.copy()
    dftemp1.dropna(inplace = True)

    dftempa = datatemplate[[colchk]].copy()
    dftempa.dropna(inplace = True)
    dftempa = dftempa.copy()

    dftemp[colres] = dftemp1[colchk].isin(dftempa[colchk])
    dftemp[colres].fillna("optional", inplace=True)

    datasrc[colres] = dftemp[colres]




def check_required_col(colchk,colres,datasrc,datatemplate,colchk1):

    if not colchk in datasrc.columns:
        print(colchk + " column not in source")
        return
    #this removes blank rows from the raw data column and template column and compares them.  It fills in the blank rows with "required"

    dftemp = datasrc[[colchk]].copy()

    dftemp1 = dftemp.copy()
    dftemp1.dropna(inplace = True)

    dftempa = datatemplate[[colchk1]].copy()
    dftempa.dropna(inplace = True)
    dftempa = dftempa.copy()

    dftemp[colres] = dftemp1[colchk].isin(dftempa[colchk1])
    dftemp[colres].fillna("required", inplace=True)

    datasrc[colres] = dftemp[colres]



def opt_in(df,name):
    #cleans up optin values
    df.loc[df[name] == 'Yes', name] = 'Y'
    df.loc[df[name] == 'No', name] = 'N'
    df.loc[df[name].isnull(), name] = 'U'
    return df


def check_exists(colname, colresult,dataframe):
    dataframe.loc[dataframe[colname].isnull(), colresult] = 'required'
    dataframe.loc[dataframe[colname].notnull(), colresult] = 'TRUE'

def check_phone_col(colchk,colres,datasrc):
    #this checks the optional columns filtering out the blanks

    if not colchk in datasrc.columns:
        print(colchk + " column not in source")
        return

    datasrc.loc[datasrc[colchk].isnull(), colres] = 'set OIP N'



#def zip_check(colchk,colres,datasrc):

import pandas as pd
df = pd.read_excel(r"C:\test\testrawdata.xlsx")
df_zip = pd.read_excel(r'C:\test\ZIP_Locale_Detail.xls')

df_zip_temp = pd.DataFrame()
df_zip_temp['zip5'] = df_zip['PHYSICAL ZIP'].astype(str).replace(r'\.0$', '', regex=True).str.zfill(5)
df_zip_temp['zip10'] = df_zip['PHYSICAL ZIP'].astype(str).replace(r'\.0$', '', regex=True).str.zfill(5) + '-' + (df_zip['PHYSICAL ZIP 4'].astype(str).replace(r'\.0$', '', regex=True).str.zfill(5))

#if not colchk in datasrc.columns:
#    print(colchk + " column not in source")
#    return
# this removes blank rows from the raw data column and template column and compares them.  It fills in the blank rows with "required"

#dftemp = datasrc[[colchk]].copy()

#dftemp1 = dftemp.copy()
#dftemp1.dropna(inplace=True)

#dftempa = datatemplate[[colchk1]].copy()
#dftempa.dropna(inplace=True)
#dftempa = dftempa.copy()

df['Zip Status'] = df['Zip'].isin(df_zip_temp())
#dftemp[colres].fillna("required", inplace=True)

print(df['Zip Status'])