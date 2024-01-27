import pandas as pd
from datetime import datetime, timedelta
import openpyxl
import os
import warnings

with warnings.catch_warnings():
    warnings.filterwarnings("ignore")


weights_file_path = 'Data\index_data\weights.xlsx'
daily_data_folder_path = 'Data\index_data\daily_data'
security_file_path = "Data\index_archieve\security.csv"


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


# form a dataframe all period stocks
def get_dataframe(start, end):

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


# calculating stocks capitalisation for date stocks
def get_date_MC(date, start):

    files = os.listdir(daily_data_folder_path)
    file_pattern = "data_{:04d}-{:02d}-{:02d}.csv" #pattern for filenames

    file_path = file_pattern.format(date.year, date.month, date.day)
    data_path = os.path.join(daily_data_folder_path, file_path)

    if file_path in files:

        stock_data = pd.read_csv(data_path, usecols=["waprice", "ticker"])

        weights_data = get_weights(start)

        df_merged = pd.merge(stock_data, weights_data, left_on='ticker', right_on='Code', how='left') # left join (left - archieve data)

        df_merged = df_merged.rename(columns={'waprice': 'P',
                                    'Number of issued shares': 'Q',
                                    'Free-float factor': 'FF',
                                    'Restricting coefficient (new)': 'RC',
                                    'Share weight in index': 'W'})

        df_merged["res"] =  df_merged["P"] * df_merged["Q"] * df_merged["FF"] * df_merged["W"] 
        df_merged.to_csv(f'index_results\index_archieve\{date}.csv')

        return df_merged["res"].sum() # MC for one date 


# form index weights for needed period
def get_weights(date):

    date_str = date.strftime("%d.%m.%Y")

    weights = pd.read_excel(weights_file_path, sheet_name=date_str, skiprows=3, usecols=['Code', 'Number of issued shares', 'Share weight in index', 'Free-float factor', 'Restricting coefficient (new)'])
    weights = weights.set_index('Code')

    return weights[:"YNDX"]


# getting Divisor
def get_date_D(date):
    data = pd.read_csv(security_file_path, skiprows=2, sep=";", encoding='latin-1', usecols=['DIVISOR', 'TRADEDATE'])
    date_str = date.strftime("%d.%m.%Y")
    divisor = data.loc[data['TRADEDATE'] == date_str, 'DIVISOR'].values[0]
    return divisor



sheets = get_valid_sheet() 
period_borders = get_period_borders_from_sheets(sheets)



df = pd.DataFrame(columns=['Date', 'MC', 'D', 'NewClose'])

for i in range (len(period_borders)-1):

    start_date = period_borders[i]
    cur_date = period_borders[i]
    end_date = period_borders[i+1]

    if start_date == datetime.strptime("2023-12-21", "%Y-%m-%d"):
        divisor = 1771049604.4106
    else:
        divisor = get_date_D(start_date)

    while cur_date < end_date:
        print(cur_date)
        MC = get_date_MC(cur_date, start_date)
        print("D: ", divisor, "  MC: ", MC)
        if MC:
            row = {'Date': cur_date, 'MC': MC, 'D': divisor, 'NewClose': MC/divisor}
            df.loc[len(df.index)] = [cur_date, MC, divisor, MC/divisor]
        cur_date += timedelta(days=1)


old_data = pd.read_csv(security_file_path, skiprows=2, sep=";", encoding='latin-1', usecols=['TRADEDATE', 'CLOSE'])

df['Date'] = pd.to_datetime(df['Date'], format="%d.%m.%Y")
old_data['TRADEDATE'] = pd.to_datetime(old_data['TRADEDATE'], format="%d.%m.%Y")
df['Date'] = pd.to_datetime(df['Date'], format="%m.%d.%Y")
old_data['TRADEDATE'] = pd.to_datetime(old_data['TRADEDATE'], format="%m.%d.%Y")

df = pd.merge(df, old_data, left_on='Date', right_on='TRADEDATE', how='left')

df.drop("TRADEDATE", axis=1).to_csv('index_results\index_calculation.csv')

print(df)





