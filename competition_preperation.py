import pandas as pd

def concat(gliederung, ctn, cc):
    if ctn < 2:
        return gliederung
    count = ctn - cc
    return f'{gliederung} {count}'

def sanitize(filename):
    # Get sheet
    df = pd.read_csv(filename, sep=';', encoding='latin-1')
    # Remove Unnamed columns
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    # Remove unessesary whitespaces
    df['gliederung'] = df['gliederung'].str.strip()
    # Count teams from same organization, age group and gender
    df['ctn'] = df.groupby(['gliederung', 'ak', 'geschlecht'])['gliederung'].transform('count')
    df['cc'] = df.groupby(['gliederung', 'ak', 'geschlecht'])['gliederung'].cumcount(ascending=False)
    # Concat team name
    df['name'] = df.apply(lambda x: concat(x['gliederung'], x['ctn'], x['cc']), axis=1)
    # Remove temporary columns
    df.drop(columns=['ctn', 'cc'], inplace=True)

    return df
