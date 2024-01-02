import pandas as pd
import openpyxl
import os


weights_file_path = 'Data\index_data\weights.xlsx'
daily_data_folder_path = 'Data\index_data\daily_data'


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

def get_dateframe(folder_path):
    files = os.listdir(folder_path)
    for file in files:
        path =  os.path.join(folder_path, file)
        cur_df = pd.read_csv(path)
        print(path)
        print(cur_df)



sheets = get_valid_sheet(weights_file_path)
get_dateframe(daily_data_folder_path)




