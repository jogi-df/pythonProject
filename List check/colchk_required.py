import pandas as pd

df_states = pd.read_excel(r'C:\test\Lead Import Template Worldwide.xlsx', sheet_name="Territory State List")
df = pd.read_excel(r"C:\test\testrawdata.xlsx")
dfa = pd.read_excel(r'C:\test\Lead Import Template Worldwide.xlsx', sheet_name="Acceptable List of Values")
df_zip = pd.read_excel(r'C:\test\ZIP_Locale_Detail.xls')

#zip code file creation

df_zip.sheet_names



def check_required_col(colchk,colres,datasrc,datatemplate,colchk1):
    #this checks the required columns filtering out the blanks

    dftemp = datasrc[[colchk]].copy()

    dftemp1 = dftemp.copy()
        #this is the source list no blanks
    dftemp1.dropna(inplace = True)

    dftempa = datatemplate[[colchk1]].copy()
        #this is the library list no blanks
    dftempa.dropna(inplace = True)
    dftempa = dftempa.copy()
    print(df_zip['ZIP CODE'])

    dftemp[colres] = dftemp1[colchk].isin(dftempa[colchk1])
    dftemp[colres].fillna("required", inplace=True)

    datasrc[colres] = dftemp[colres]

    df.to_excel(r'C:\test\working_copy.xlsx', index=False)




check_required_col('Zip','Zip Good',df,df_zip,'DELIVERY ZIPCODE')