import pandas as pd
from datetime import datetime, timedelta
import openpyxl
import os


weights_file_path = 'Data\index_data\weights.xlsx'
daily_data_folder_path = 'Data\index_data\daily_data'


# get sheets with relevant wights
def get_valid_sheet(file_path):

    workbook = openpyxl.load_workbook(file_path)
    sheets = workbook.sheetnames
    sheets.remove('help')

    last_valid_sheet_index = 0
    valid_period = True
    while valid_period:
        sheet = sheets[last_valid_sheet_index]
        worksheet = workbook[sheet]

        try:
            tab_color = worksheet.sheet_properties.tabColor.rgb
            last_valid_sheet_index += 1

        except AttributeError:
            valid_period = False
            print("STOP")

    return sheets[:last_valid_sheet_index]


# get starts and ends of periods from sheets (end of last perioв is today or last parsed date)
def get_period_borders_from_sheets(sheets):
    dates = [datetime.strptime(date, '%d.%m.%Y').date() for date in sheets] # list of starts of calculation periods 
    dates.insert(0, datetime.today().date()) # ?last parsed date
    dates = dates[::-1]
    return dates


# form a dataframe with needed stock's price data
def get_dateframe(folder_path, start, end):
    files = os.listdir(folder_path)
    file_pattern = "data_{:04d}-{:02d}-{:02d}.csv"
    cur_date = start
    df = pd.DataFrame()

    while cur_date <= end:
        data_path = file_pattern.format(cur_date.year, cur_date.month, cur_date.day)

        if data_path in files:
            data = pd.read_csv(data_path).waprice

        cur_date += timedelta(days=1)

    return df



sheets = get_valid_sheet(weights_file_path) 
period_borders = get_period_borders_from_sheets(sheets)

df = get_dateframe(daily_data_folder_path, period_borders[0], period_borders[1])
print(df)

# check every period
# for i in range (len(period_borders)-1):
#     start_date = period_borders[i]
#     end_date = period_borders[i+1]
#     df = get_dateframe(daily_data_folder_path, start_date, end_date)




