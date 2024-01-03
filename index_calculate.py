import pandas as pd
from datetime import datetime
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
    for file in files:
        path =  os.path.join(folder_path, file)
        cur_df = pd.read_csv(path)
        cur_df.waprice
        print(path)
        print(cur_df)



sheets = get_valid_sheet(weights_file_path) 
period_borders = get_period_borders_from_sheets(sheets)

print(period_borders)


# check every period
for i in range (len(period_borders)-1):
    start_date = period_borders[i]
    end_date = period_borders[i+1]
    df = get_dateframe(daily_data_folder_path, start_date, end_date)




