import pandas as pd
from datetime import datetime, timedelta
import openpyxl
import os
import warnings

with warnings.catch_warnings():
    warnings.filterwarnings("ignore")


weights_file_path = 'Data\index_data\weights.xlsx'
daily_data_folder_path = 'Data\index_data\daily_data'


# get sheets with relevant wights
def get_valid_sheet():

    # get all sheets
    workbook = openpyxl.load_workbook(weights_file_path)
    sheets = workbook.sheetnames
    sheets.remove('help')

    # find index of last valid sheet in excel file
    last_valid_sheet_index = 0
    valid_period = True
    while valid_period:
        sheet = sheets[last_valid_sheet_index]
        worksheet = workbook[sheet]

        # find all sheets that are highlighted green
        try:
            tab_color = worksheet.sheet_properties.tabColor.rgb
            last_valid_sheet_index += 1

        except AttributeError:
            valid_period = False

    return sheets[:last_valid_sheet_index]


# get starts and ends of periods from sheets (end of last period is today or ?last parsed date)
def get_period_borders_from_sheets(sheets):
    dates = [datetime.strptime(date, '%d.%m.%Y').date() for date in sheets] # list of starts of calculation periods 
    dates.insert(0, datetime.today().date()) 
    dates = dates[::-1]
    return dates


# form a dataframe with needed stock's price data
def get_dateframe(start, end):

    files = os.listdir(daily_data_folder_path)
    file_pattern = "data_{:04d}-{:02d}-{:02d}.csv" #pattern for filenames
    cur_date = start
    first_fill = True # for dataframe initing

    # getting data from all needed files by date
    while cur_date <= end:

        file_path = file_pattern.format(cur_date.year, cur_date.month, cur_date.day)
        data_path = os.path.join(daily_data_folder_path, file_path)

        # get data if file exists
        if file_path in files:
            # reading data
            data = pd.read_csv(data_path, usecols=["tradedate", "waprice", "secids"])

            # form new row
            row_to_add = {"date": data.tradedate[0]}
            row_to_add.update(dict(zip(data.secids, data.waprice))) 

            # init a dataframe firstly
            if first_fill:
                df = pd.DataFrame(columns=["date"]+data.secids.tolist())
                first_fill = False  

            # add new row
            df = pd.concat([df, pd.DataFrame(row_to_add, index=[0])]) 

        cur_date += timedelta(days=1)

    df.set_index('date', inplace=True)
    return df


# form index weights for needed period
def get_weights(date):

    weights = pd.read_excel(weights_file_path, sheet_name=date, skiprows=3, usecols=['Code', 'Weight new', 'Free-float factor', 'Restricting coefficient (new)'])
    weights = weights.set_index('Code')

    return weights[:"YNDX"]



sheets = get_valid_sheet() 
period_borders = get_period_borders_from_sheets(sheets)

df = get_dateframe(period_borders[0], period_borders[1])
weights = get_weights(sheets[0])



# check every period
# for i in range (len(period_borders)-1):
#     start_date = period_borders[i]
#     end_date = period_borders[i+1]
#     df = get_dateframe(daily_data_folder_path, start_date, end_date)




