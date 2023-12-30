import pandas as pd
import openpyxl

weights_file_path = 'Data\index_daily_data\weights.xlsx'


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


sheets = get_valid_sheet(weights_file_path)

for sheet in sheets:
    weights = pd.read_excel(weights_file_path, sheet_name=sheet, skiprows=3)
    print(weights.columns)
