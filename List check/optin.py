
def opt_in(df,name):
    df.loc[df[name] == 'Yes', name = 'Y'
    df.loc[df[name] == 'No', name] = 'N'
    df.loc[df[name].isnull(), name] = 'U'
    return df






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