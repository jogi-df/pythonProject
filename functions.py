
def check_optional_col(colchk,colres,datasrc,datatemplate):          #this checks the required columns filtering out the blanks

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



def opt_in(df,name):                            #cleans up optin values
    df.loc[df[name] == 'Yes', name] = 'Y'
    df.loc[df[name] == 'No', name] = 'N'
    df.loc[df[name].isnull(), name] = 'U'
    return df
