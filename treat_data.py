# Importando bibliotecas
from datetime import datetime
from datetime import timedelta
import xlrd
import numpy as np
import pandas as pd
import os
import seaborn as sns

# ================================================================
# PROGRAM CONSTANTS
# PATHs
PROJECT_PATH = os.path.dirname(os.path.abspath(__file__))
TREATED_PATH = os.path.join(PROJECT_PATH, r'data\treated')
ELECTRIC_PATH = os.path.join(PROJECT_PATH, r'data\raw\Consumption_data')
WIND_PATH = os.path.join(PROJECT_PATH, r'data\raw\Weather_data\Wind')
WET_TEMP_PATH = os.path.join(PROJECT_PATH, r'data\raw\Weather_data\Wet_bulb_temperature')
DRY_TEMP_PATH = os.path.join(PROJECT_PATH, r'data\raw\Weather_data\Dry_bulb_temperature')
HUMID_PATH = os.path.join(PROJECT_PATH, r'data\raw\Weather_data\Humidity')
RAD_PATH = os.path.join(PROJECT_PATH, r'data\raw\Weather_data\Radiation')
PRESS_PATH = os.path.join(PROJECT_PATH, r'data\raw\Weather_data\Pressure')

# PARAMETERS
SAVE_TREATED_DATA = True  # Toggle this true if you want to save the treated data to the treated_path

# ================================================================
# DEFINING AUXILIARY FUNCTIONS
def excel_folder_to_df(data_abs_path, treat_function, var_name=None):
    '''
    The function piles all excel data in a directory into a single dataframe
    Given the funcion that treats the excels in that folder
    Input:
        data_abs_path: absolute path of the folder containing all the Excel
        treat_function: function that treats the file of the fomat: function(filename)
        var_name (defalut: None): string of the variable name to rename
    Output:
        pd.DataFrame
    '''
    df_all = pd.DataFrame()
    
    excel_list = [f for f in os.listdir(data_abs_path) if os.path.isfile(os.path.join(data_abs_path, f))]
    
    for excel in excel_list:
        file_name = os.path.join(data_abs_path, excel)
        print("Treating the file: ", file_name)
        df = treat_function(file_name)
        df_all = df_all.append(df)
    
    if var_name:
        df_all.rename(columns={'value': var_name}, inplace=True)
    
    return df_all.sort_index()


def parse_pt_date(date_string):
    '''
    Parses a date-time string in the format 'JANEIRO de 2007' and returns a string in the isoformat '2007-01'
    Or 'JAN. 2007'
    '''
    MONTHS = {'jan': 1, 'fev': 2, 'mar': 3, 'abr': 4,  'mai': 5,  'jun': 6,
          'jul': 7, 'ago': 8, 'set': 9, 'out': 10, 'nov': 11, 'dez': 12}
    
    FULL_MONTHS = {'janeiro': 1,  'fevereiro': 2, u'março': 3,    'abril': 4,
                'maio': 5,     'junho': 6,     'julho': 7,     'agosto': 8,
                'setembro': 9, 'outubro': 10,  'novembro': 11, 'dezembro': 12}
    
    try:
        date_string = date_string.replace('-', '')
        date_string = date_string.replace('.', '')
        date_info = date_string.lower().split()
        
        if 'de' in date_info:
            date_info.remove('de')
        
        month_pt = date_info[0]
        year = date_info[1]
        
        if len(month_pt) == 3:
            month = MONTHS[month_pt]
        else:
            month = FULL_MONTHS[month_pt]
        
        date_iso = '{}-{:02d}'.format(int(year), int(month)) # Adding to a string on YYYY-MM format
        
        return date_iso
    except Exception as e:
        print("Erro ao tratar o ano e mês: ", date_string)


def datestr_to_datetime(year_month, day, hour):
    if int(hour) != 24:
        date_str = year_month + "-{:02d}, {:02d}".format(int(day), int(hour))
        
        return pd.to_datetime(date_str, format='%Y-%m-%d, %H')
    else:
        hour = 0
        date_str = year_month + "-{:02d}, {:02d}".format(int(day), int(hour))
        return pd.to_datetime(date_str, format='%Y-%m-%d, %H') + timedelta(days=1)
    
def treat_wind(filename):
    '''
    Reads the Wind Excel file and returns a DataFrame with the wind velocity and direction on the columns
    '''
    
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


def treat_weather(filename):
    '''
    Reads the Excel files of the "Dry bulb temperature", "Wet bulb temperature", "Humidity" and "Pressure"
    Returns a DataFrame with the variable on the columns and the date on the index
    '''
    
    # Reading the first cells of the Excel workbooks that contain the year and the month as a string
    book = xlrd.open_workbook(filename)
    
    # Each sheet has the data of a month. Listing the sheet names
    sheet_names = book.sheet_names()
    
    # Iterating for every month in sheet_names and appending to the df_year:
    df_year = pd.DataFrame()
    
    for month in sheet_names:
        # Reading the second cell where the Year and the Months are
        month_sheet = book.sheet_by_name(month)

        if month_sheet.cell_value(rowx=0, colx=1):
            year_month = month_sheet.cell_value(rowx=0, colx=1)
        else:
            print("Error reading the year-month:", month_sheet.cell_value(rowx=0, colx=1))
        
        # Treating the year and month that are in Portuguese
        year_month = parse_pt_date(year_month)
        
        # Reading the current sheet of the month. The hours are in the rows and days in the columns
        df_month_raw = pd.read_excel(filename,
                                     sheet_name=month,
                                     header=2,
                                     usecols="A:AF",
                                     nrows=24)
        
        df_treated = df_month_raw.copy()
        
        # Removing the days that have all null values (e.g.: February has no day 31)
        df_treated.dropna(how="all", axis = "columns", inplace=True)
        
        # Removing the null hours (e.g.: The wet bulb temperature has no data from 1h to 6h)
        df_treated = df_treated.loc[~(df_treated['H/D'].isna())]
        
        df_treated.rename(columns={'H/D': 'hour'}, inplace=True)
        
        # Unpivoting the days that are in the columns to the rows
        df_treated = df_treated.melt(id_vars=['hour'])
        df_treated.rename(columns={'variable': 'day'} , inplace=True)
        
        # Treating the date to the datetime format
        df_treated['date'] = df_treated.apply(lambda row: datestr_to_datetime(year_month, row['day'], row['hour']),
                                              axis=1)
        
        # Reordering the columns for better visualization
        df_month = df_treated[['date', 'value']].set_index('date')

        df_year = df_year.append(df_month)
        
    return df_year
    

def treat_radiation(filename):
    # Reading the Excel workbook to sweep through their sheets
    book = xlrd.open_workbook(filename)
    
    # Each sheet has the data of a month. Listing the sheet names
    sheet_names = book.sheet_names()
    
    # Iterating for every month in sheet_names and appending to the df_year:
    df_all_rad = pd.DataFrame()
    
    for year in sheet_names:
        # Reading the current sheet of the year. The days are in the rows and the months in the columns
        df_year_rad_raw = pd.read_excel(filename,
                                        sheet_name=year,
                                        header=5,
                                        usecols="B:N",
                                        nrows=32)
        
        df_treated = df_year_rad_raw.copy()
        
        # Renaming the days column
        df_treated.rename(columns={'DIA / MÊS': 'day'}, inplace=True)
        
        df_treated.dropna(subset=['day'], inplace=True)
        
        df_treated = df_treated.melt(id_vars=['day'])
        
        df_treated.rename(columns={'variable': 'month',
                                   'value': 'radiation'},
                          inplace=True)
        
        df_treated.dropna(subset=['radiation'], inplace=True)  # Removing days without data (e.g. Feb 29th)
        
        df_treated['year_month'] = df_treated.apply(lambda row: parse_pt_date(row['month'] + ' ' + str(int(year)) ),
                                                    axis=1)
        
        df_treated['date'] = pd.to_datetime(df_treated['year_month'] + '-' + df_treated['day'].astype(int).astype(str),
                                            format="%Y-%m-%d")
        
        df_rad = df_treated[['date', 'radiation']].set_index('date')
        
        df_all_rad = df_all_rad.append(df_rad)
    
    return df_all_rad


def treat_consumption(filename):
    df_consumption = pd.read_excel(filename,parse_dates=['Data'])
    
    # treating column names
    df_consumption.rename({'Data': 'date',
                           'Reposição de demanda': 'demand_replacement',
                           'Intervalo reativo': 'reactive_delay',
                           'Posto': 'station'},
                          axis='columns',
                          inplace=True)
    
    df_consumption.columns = df_consumption.columns.str.strip().str.lower().str.replace(' ', '_', regex=False).str.replace('[', '', regex=False).str.replace(']', '', regex=False)
    
    # setting date as index
    df_consumption.set_index('date', drop = True, inplace=True)
    
    return df_consumption


# ================================================================
# START OF THE PROGRAM

# Treating all the weather files
df_all_wind = excel_folder_to_df(WIND_PATH, treat_wind)  # Piling all the wind files into one df
df_all_wet_temp = excel_folder_to_df(WET_TEMP_PATH, treat_weather, 'wet_temp')  # Piling all the wet bulb temperature files into one df
df_all_dry_temp = excel_folder_to_df(DRY_TEMP_PATH, treat_weather, 'dry_temp')  # Piling all the dry bulb temperature files into one df
df_all_humid = excel_folder_to_df(HUMID_PATH, treat_weather, 'humidity')  # Piling all the wet Humidity files into one df
df_all_pres = excel_folder_to_df(PRESS_PATH, treat_weather, 'pressure')  # Piling all the Pressure files into one df
df_all_rad = excel_folder_to_df(RAD_PATH, treat_radiation)  # Piling all the Radiation files into one df
df_all_consumpt = excel_folder_to_df(ELECTRIC_PATH, treat_consumption)  # Piling all the Radiation files into one df

# Saving the treated data:
if SAVE_TREATED_DATA:
    print("Saving the treated data...")
    df_all_wind.to_csv(os.path.join(TREATED_PATH, 'wind.csv'))
    df_all_dry_temp.to_csv(os.path.join(TREATED_PATH, 'dry_temperature.csv'))
    df_all_wet_temp.to_csv(os.path.join(TREATED_PATH, 'wet_temperature.csv'))
    df_all_humid.to_csv(os.path.join(TREATED_PATH, 'humidity.csv'))
    df_all_pres.to_csv(os.path.join(TREATED_PATH, 'pressure.csv'))
    df_all_rad.to_csv(os.path.join(TREATED_PATH, 'radiation.csv'))
    df_all_consumpt.to_csv(os.path.join(TREATED_PATH, 'electric_consumption.csv'))