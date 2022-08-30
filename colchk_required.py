import pandas as pd


def check_required_col(colchk,colres,datasrc,datatemplate):          #this checks the required columns filtering out the blanks

    #df = pd.read_excel(r"C:\test\testrawdata.xlsx")
    #dfa = pd.read_excel(r'C:\test\Lead Import Template Worldwide.xlsx', sheet_name="Acceptable List of Values")

    dftemp = datasrc[[colchk]].copy()

    dftemp1 = dftemp.copy()
    dftemp1.dropna(inplace = True)

    dftempa = datatemplate[[colchk]].copy()
    dftempa.dropna(inplace = True)
    dftempa = dftempa.copy()

    dftemp['correct_in_list'] = dftemp1[colchk].isin(dftempa[colchk])
    dftemp['correct_in_list'].fillna("required", inplace=True)

    datasrc['correct_in_list'] = dftemp['correct_in_list']

    #df.to_excel(r'C:\test\working_copy.xlsx', index=False)



check_required_col('Estimated Number of Units','Est # Units')
