#. /C/ProgramData/Miniconda3/etc/profile.d/conda.sh
# C:\\ProgramData\\Miniconda3\\condabin\\activate.bat C:\\ProgramData\\Miniconda3\\envs\\cd-python-tcc

# Importando bibliotecas
from datetime import datetime
import xlrd
import numpy as np
import pandas as pd
import os

def parse_pt_date(date_string):
    '''
    Parses a date-time string in the format 'JANEIRO de 2007' and returns a string in the isoformat '2007-01'
    '''
    MONTHS = {'jan': 1, 'fev': 2, 'mar': 3, 'abr': 4,  'mai': 5,  'jun': 6,
          'jul': 7, 'ago': 8, 'set': 9, 'out': 10, 'nov': 11, 'dez': 12}
    
    FULL_MONTHS = {'janeiro': 1,  'fevereiro': 2, u'março': 3,    'abril': 4,
                'maio': 5,     'junho': 6,     'julho': 7,     'agosto': 8,
                'setembro': 9, 'outubro': 10,  'novembro': 11, 'dezembro': 12}
    
    try:
        date_string = date_string.replace('-', '')
        date_info = date_string.lower().split()
        date_info.remove('de')
        
        month_pt = date_info[0]
        year = date_info[1]
        
        if len(month_pt) == 3:
            month = MONTHS[month_pt]
        else:
            month = FULL_MONTHS[month_pt]
        
        date_iso = '{}-{:02d}'.format(year, int(month)) # Adding to a string on YYYY-MM format
        
        return date_iso
    except Exception as e:
        print("Erro ao tratar o ano e mês: ", date_string)
    
def treat_wind(filename):
    '''
    Reads the Wind Excel file and returns a DataFrame with the wind velocity and direction on the columns
    '''
    print("Treating the file: ", filename)
    
    # Reading the first cells of the Excel workbooks that contain the year and the month as a string
    book = xlrd.open_workbook(filename)
    sheet = book.sheet_by_index(0)
    
    cell0 = sheet.cell_value(rowx=0, colx=0)
    cell1 = sheet.cell_value(rowx=0, colx=1)
    cell2 = sheet.cell_value(rowx=0, colx=2)
    
    if type(cell2) != str:
        if len(cell0) != 0 and len(cell1) != 0 and cell2 != None:
            # Files that have the month in the first cell and the year in the third
            month = cell0
            year = cell2
            year_month = month + " de " + str(int(year))
    else:
        if len(cell0) != 0 and len(cell1) == 0 and len(cell2) == 0:
            year_month = cell0  # Files that have both the year and month in the first cell
        elif len(cell0) == 0 and len(cell1) != 0 and len(cell2) == 0:
            year_month = cell1  # Files that have both the year and month in the second cell
        elif len(cell0) == 0 and len(cell1) == 0 and len(cell2) != 0:
            year_month = cell2  # Files that have both the year and month in the third cell
        elif len(cell0) != 0 and len(cell1) != 0 and len(cell2) == 0:
            # Files that have both the month in the first cell and the year in the second
            month = sheet.cell_value(rowx=0, colx=0)
            year = sheet.cell_value(rowx=0, colx=1)
            year_month = month + " " + year
        elif len(cell0) != 0 and len(cell1) != 0 and cell2 != None:
            # Files that have the month in the first cell and the year in the third
            month = sheet.cell_value(rowx=0, colx=0)
            year = sheet.cell_value(rowx=0, colx=2)
            year_month = month + " de " + str(year)
        else:
            print('Error reading the year_month. Cells: ', cell0, len(cell0), cell1, len(cel1), cell2, len(cell2))
            
    year_month = parse_pt_date(year_month)
    
    df_raw_wind = pd.read_excel(filename, header=7, usecols="A:AW", nrows=36)
    
    df_treated = df_raw_wind.copy()
    
    # Cleaning the df
    df_treated = df_treated.loc[(df_treated['Hs.'] != 'DIA') & (df_treated['Hs.'] !='Hs.') & ~(df_treated['Hs.'].isna())]
    df_treated.rename(columns={'Hs.': 'day'}, inplace=True)
    
    df_treated.set_index('day', drop = True, inplace=True)
    
    hours = np.arange(0,24)
    vars = ["wind_dir", "wind_vel"]
    
    df_treated.columns = pd.MultiIndex.from_product([hours, vars], names=['hour', 'vars'])
    
    df_wind = df_treated.stack(0)
    
    df_wind.reset_index(inplace=True)
    
    # Adding the column 'date' combining the year_month and the columns 'day' and 'hour'
    df_wind['date'] = df_wind.apply(lambda row: year_month + "-{:02d}, {:02d}".format(row['day'], row['hour']),
                                    axis=1)
    
    df_wind['date'] = pd.to_datetime(df_wind['date'], format="%Y-%m-%d, %H")
    
    df_wind.drop(columns = ['day','hour'], inplace=True)
    
    # Reordering the columns for better visualization
    df_wind = df_wind[['date', 'wind_vel', 'wind_dir']]

    df_wind.set_index('date', inplace=True)
    
    return df_wind

wind_raw_path = r'D:\Bibliotecas\Documents\Poli\TCC\Programa\Energy_consumption_forecasting\Data\Weather_data\raw\Wind'
TREATED_PATH = r'D:\Bibliotecas\Documents\Poli\TCC\Programa\Energy_consumption_forecasting\Data\Weather_data\treated'

excel_list = [f for f in os.listdir(wind_raw_path) if os.path.isfile(os.path.join(wind_raw_path, f))]

df_all_wind = pd.DataFrame()

for excl in excel_list:
    df_wind = treat_wind(os.path.join(wind_raw_path, excl))
    df_all_wind = df_all_wind.append(df_wind)

df_all_wind.sort_index().to_excel(os.path.join(TREATED_PATH, 'wind.xls'))


# df_wind = treat_wind(r'D:\Bibliotecas\Documents\Poli\TCC\Programa\Energy_consumption_forecasting\Data\Weather_data\raw\Wind\04V2006b F S.xls')
