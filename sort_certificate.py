import re
import pandas as pd

age_groups = ['AK 10', 'AK 12', 'AK 13/14', 'AK 15/16', 'AK 17/18', 'AK offen', 'AK Senioren',
              'AkW 13/14', 'AkW 15/16', 'AkW 17/18', 'AkW offen', 'AkW Senioren']

def sort(filename):
    # Get sheet
    df = pd.read_excel(filename, sheet_name='Seriendruck', index_col=0)

    # Fix names when something is wrong
    regex_pat = re.compile(r'ak', flags=re.IGNORECASE)
    df['Altersklasse'].replace(to_replace=regex_pat, value='AK', regex=True, inplace=True)
    regex_pat = re.compile(r'akw', flags=re.IGNORECASE)
    df['Altersklasse'].replace(to_replace=regex_pat, value='AkW', regex=True, inplace=True)

    # Predefine category sort
    df['Altersklasse'] = pd.Categorical(df['Altersklasse'], age_groups)
    # Sort values
    df.sort_values(by=['Altersklasse', 'Geschlecht', 'Platz'], ascending=[True, False, False], inplace=True)

    return df
