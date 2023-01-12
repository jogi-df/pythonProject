import numpy as np
from uszipcode import SearchEngine

search = SearchEngine()
# United States of America Python Dictionary to translate States,
# Districts & Territories to Two-Letter codes and vice versa.
#
# also added in Canadian territories.  yes i know they're not us states.

state_to_abbrev = {
    "Alabama": "AL",
    "Alaska": "AK",
    "Arizona": "AZ",
    "Arkansas": "AR",
    "California": "CA",
    "Colorado": "CO",
    "Connecticut": "CT",
    "Delaware": "DE",
    "Florida": "FL",
    "Georgia": "GA",
    "Hawaii": "HI",
    "Idaho": "ID",
    "Illinois": "IL",
    "Indiana": "IN",
    "Iowa": "IA",
    "Kansas": "KS",
    "Kentucky": "KY",
    "Louisiana": "LA",
    "Maine": "ME",
    "Maryland": "MD",
    "Massachusetts": "MA",
    "Michigan": "MI",
    "Minnesota": "MN",
    "Mississippi": "MS",
    "Missouri": "MO",
    "Montana": "MT",
    "Nebraska": "NE",
    "Nevada": "NV",
    "New Hampshire": "NH",
    "New Jersey": "NJ",
    "New Mexico": "NM",
    "New York": "NY",
    "North Carolina": "NC",
    "North Dakota": "ND",
    "Ohio": "OH",
    "Oklahoma": "OK",
    "Oregon": "OR",
    "Pennsylvania": "PA",
    "Rhode Island": "RI",
    "South Carolina": "SC",
    "South Dakota": "SD",
    "Tennessee": "TN",
    "Texas": "TX",
    "Utah": "UT",
    "Vermont": "VT",
    "Virginia": "VA",
    "Washington": "WA",
    "West Virginia": "WV",
    "Wisconsin": "WI",
    "Wyoming": "WY",
    "District Of Columbia": "DC",
    "American Samoa": "AS",
    "Guam": "GU",
    "Northern Mariana Islands": "MP",
    "Puerto Rico": "PR",
    "United States Minor Outlying Islands": "UM",
    "U.S. Virgin Islands": "VI",
    'Alberta': 'AB',
    'British Columbia':'BC',
    'Manitoba':'MB',
    'New Brunswick': 'NB',
    'Newfoundland and Labrador': 'NL',
    'Northwest Territories': 'NT',
    'Nova Scotia': 'NS',
    'Nunavut': 'NU',
    'Ontario': 'ON',
    'Prince Edward Island': 'PE',
    'Quebec': 'QC',
    'Saskatchewan': 'SK',
    'Yukon': 'YT'
}

# invert the dictionary
# abbrev_to_us_state = dict(map(reversed, us_state_to_abbrev.items()))

def state_check(colchk, datasrc):       #replace state names with abbreviation
    datasrc[colchk] = datasrc[colchk].replace(state_to_abbrev)
    return

def check_optional_col(colchk,colres,datasrc,datatemplate):
    #this checks the optional columns filtering out the blanks

    if not colchk in datasrc.columns:
        print(colchk + " column not in source")
        return

    datasrc.loc[datasrc[colchk].isnull(), colres] = 'optional'

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
        print(colchk + " column not in source and is REQUIRED")
        return
    #this removes blank rows from the raw data column and template column and compares them.  It fills in the blank rows with "required"

    dftemp = datasrc[[colchk]].copy()

    dftemp1 = dftemp.copy()
    dftemp1.dropna(inplace = True)

    dftempa = datatemplate[[colchk1]].copy()
    dftempa.dropna(inplace = True)
    dftempa = dftempa.copy()

    dftemp[colres] = dftemp1[colchk].isin(dftempa[colchk1])
    dftemp[colres].fillna("required", inplace = True)

    datasrc[colres] = dftemp[colres]




def opt_in_check(df,name):
    #cleans up optin values
    df.loc[df[name] == 'Yes', name] = 'Y'
    df.loc[df[name] == 'No', name] = 'N'
    df.loc[df[name].isnull(), name] = 'U'
    return df

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

    datasrc[colchk] = datasrc[colchk].astype(str)
#    datasrc[colres] = datasrc[colchk].str.match("^[+]?[1]?(\+\d{1,2}\s)?\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}$")
    datasrc[colres] = datasrc[colchk].str.match("^[+]?[1]?[-| |.]?[(| ]?\d{3}[)]?[-| |.]?\d{3}[-| |.]?\d{4}")
    datasrc[colchk] = datasrc[colchk].replace('nan', np.nan)
#    datasrc.loc[datasrc[colchk].isnull(), colres] = 'set OIP N'


def whitespace_remover(dataframe):
    # iterating over the columns
    for i in dataframe.columns:

        # checking datatype of each columns
        if dataframe[i].dtype == 'object':

            # applying strip function on column
            dataframe[i] = dataframe[i].map(str.strip)
        else:

            # if condn. is False then it will do nothing.
            pass


def zip_city(x):
    city = search.by_zipcode(x).major_city
    print(city)
    return city if city else 'None'

def zip_state(x):
    state = search.by_zipcode(x).state
    print(state)
    return state if state else 'None'

#def zip_check(colchk,colres,datasrc):

import pandas as pd

#df = pd.read_excel(r"C:\test\testrawdata.xlsx")
#df_zip = pd.read_excel(r'C:\test\ZIP_Locale_Detail.xls')

#df_zip_temp = pd.DataFrame()
#dfzt = pd.DataFrame()

#df_zip = pd.read_excel(r'C:\test\ZIP_Locale_Detail.xls', sheet_name='ZIP_DETAIL')
#df_zip_temp['zip5'] = df_zip.iloc[:,4].astype(str).replace(r'\.0$', '', regex=True).str.zfill(5)
#df_zip_temp['zip10'] = df_zip.iloc[:,9].astype(str).replace(r'\.0$', '', regex=True).str.zfill(5) + '-' + (df_zip.iloc[:,10].astype(str).replace(r'\.0$', '', regex=True).str.zfill(5))


#df_zip = pd.read_excel(r'C:\test\ZIP_Locale_Detail.xls', sheet_name='Unique_ZIP_DETAIL')
#dfzt['zip5'] = df_zip.iloc[:,4].astype(str).replace(r'\.0$', '', regex=True).str.zfill(5)
#print(dfzt['zip5'])

#df_zip_temp.concat(df_zip['zip5'], dfzt['zip5'])
#df_zip_temp['zip10'] = df_zip.concat([df_zip.iloc[:,9].astype(str).replace(r'\.0$', '', regex=True).str.zfill(5) + '-' + (df_zip.iloc[:,10].astype(str).replace(r'\.0$', '', regex=True).str.zfill(5)), df_zip_temp['zip10']])

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
#print(df['Zip'])
#print(df_zip_temp)

#df_zip_temp1 = (df_zip_temp.to_numpy().ravel().tolist())
#print(df_zip_temp1)
#df['Zip Status'] = df['Zip'].astype(str).isin(df_zip_temp['zip_mail'])
#print(df['Zip'].astype(str).isin(df_zip_temp1))
#print(df['Zip'].astype(str).isin(df_zip_temp).any().any())
#print(df['Zip'].astype(str).isin(df_zip_temp.to_numpy().ravel().tolist()))
#dftemp[colres].fillna("required", inplace=True)

#print(df['Zip Status'])