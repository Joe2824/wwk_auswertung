import pandas as pd

# Debugmodus
DEBUG = False
# Subdrahiere nicht angetretene Mannschaften in Gesamtwertung
SUBTRACT_NOT_STARTED = True


def calculate_function(filename):
    df = pd.read_excel(filename, sheet_name=None, skiprows=[1])

    wave_table = pd.DataFrame()
    rescure_table = pd.DataFrame()

    for name, sheet in df.items():
        sheet['Altersklasse'] = name
        sheet = sheet.rename(columns=lambda x: x.split('\n')[-1])
        if name[:1].upper() == 'A':
            sheet = clear_sheet(sheet)
            sheet = count_points(sheet)
            if name[:3].upper() == 'AKW':
                wave_table = pd.concat([wave_table, sheet])
            else:
                rescure_table = pd.concat([rescure_table, sheet])

    print('Rettungswettkampf \n\n', rescure_table) if DEBUG else 0
    print('Wellenwettkmapf \n\n', wave_table) if DEBUG else 0

    rescure_table_detail = rescure_table.copy()
    wave_table_detail = wave_table.copy()
    rescure_table = merge_values(rescure_table)
    wave_table = merge_values(wave_table)

    return wave_table, rescure_table, wave_table_detail, rescure_table_detail


def merge_values(sheet):
    sheet = sheet.groupby('Gliederung').agg({'Punkte': 'sum'}).sort_values(by='Punkte', ascending=False).reset_index()
    sheet.index += 1
    return sheet


def clear_sheet(sheet):
    sheet = sheet[['Pl', 'Gliederung', 'Altersklasse']]
    return sheet


def count_points(sheet):
    # Anzahl Reihen minus NaN Zellen in Spalte
    count_row = sheet.shape[0] - (sheet.isnull().sum() if SUBTRACT_NOT_STARTED else 0)
    # Berechnet Punktzahl
    sheet['score'] = (count_row+1 - sheet.drop(columns=['Gliederung', 'Altersklasse'])).sum(axis=1)
    # Zähle Erste Plätze
    sheet['extra_points'] = (sheet.drop(columns='score') ==1).sum(axis=1)
    # Summiere Punkzahl und anzahl erste Plätze
    sheet['Punkte'] = sheet.loc[:,['score','extra_points']].sum(axis=1)
    return sheet
