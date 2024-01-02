import requests
from datetime import datetime, timedelta
from pprint import pprint
import pandas as pd
import os

cookies = {
    '_ym_uid': '1682513184127487068',
    '_ym_d': '1698584361',
    '_ga_1RFNGQ1758': 'GS1.1.1702731461.9.0.1702731479.42.0.0',
    'tmr_lvid': '7a3ea67c015dc781a31a735757421fc3',
    'tmr_lvidTS': '1703082671793',
    'MPTrack': '2ea106bf.60cf33d6bd2d2',
    'MicexTrackID': '90f7fa46.60d061559ee07',
    '_ga_9HDREVDWE7': 'GS1.1.1703416223.7.1.1703416225.58.0.0',
    '_ga': 'GA1.2.41125269.1698584361',
    '_gid': 'GA1.2.1773643554.1703416226',
    '_gat_UA-27780661-1': '1',
    '_ym_isad': '2',
    'MicexPassportCert': 'GfbDXnESx_xTDYQU3SCyogUAAADIAxm5MFyG_LDyGs6wAn_RwSjQhR-CCvl7tJyBz0D2_U1vCUN1LhmbnFhNT__fZk0rAJDsX3j1qHxJTm0h1ygOoXVAylqg73UOobXN6NZgrCBnL5vfltVwPekUv5R4Dtpu6-sgsMtPtNv23RDkrX5KGADQqkt4Hhp-r5brDudMWGmQf4W00',
}

headers = {
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'ru,en;q=0.9,hu;q=0.8,es;q=0.7',
    'Cache-Control': 'no-cache',
    'Connection': 'keep-alive',
    # 'Cookie': '_ym_uid=1682513184127487068; _ym_d=1698584361; _ga_1RFNGQ1758=GS1.1.1702731461.9.0.1702731479.42.0.0; tmr_lvid=7a3ea67c015dc781a31a735757421fc3; tmr_lvidTS=1703082671793; MPTrack=2ea106bf.60cf33d6bd2d2; MicexTrackID=90f7fa46.60d061559ee07; _ga_9HDREVDWE7=GS1.1.1703416223.7.1.1703416225.58.0.0; _ga=GA1.2.41125269.1698584361; _gid=GA1.2.1773643554.1703416226; _gat_UA-27780661-1=1; _ym_isad=2; MicexPassportCert=GfbDXnESx_xTDYQU3SCyogUAAADIAxm5MFyG_LDyGs6wAn_RwSjQhR-CCvl7tJyBz0D2_U1vCUN1LhmbnFhNT__fZk0rAJDsX3j1qHxJTm0h1ygOoXVAylqg73UOobXN6NZgrCBnL5vfltVwPekUv5R4Dtpu6-sgsMtPtNv23RDkrX5KGADQqkt4Hhp-r5brDudMWGmQf4W00',
    'Origin': 'https://www.moex.com',
    'Pragma': 'no-cache',
    'Referer': 'https://www.moex.com/',
    'Sec-Fetch-Dest': 'empty',
    'Sec-Fetch-Mode': 'cors',
    'Sec-Fetch-Site': 'same-site',
    'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 13_2_3 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/13.0.3 Mobile/15E148 Safari/604.1',
}

params = [
    ('iss.meta', 'off'),
    ('iss.json', 'extended'),
    ('start', '0'),
    ('limit', '100000'),
    ('lang', 'ru'),
    # ('date', '2023-05-22'),
    ('nocache-guid', '25732139-cb65-4fb9-a82b-097c55e9747e'),
    ('iss.meta', 'off'),
    ('iss.json', 'extended'),
    ('callback', 'JSON_CALLBACK'),
    ('lang', 'ru'),
]

cur_date = datetime(2013, 9, 2)
end_date = datetime.today() - timedelta(days=2)

data_files = os.listdir("Data/index_data/daily_data")


def get_response(date):
    params.append(('date', date.strftime("%Y-%m-%d")))
    response = requests.get(
        'https://iss.moex.com/iss/statistics/engines/stock/markets/index/analytics/IMOEX.jsonp',
        params=params,
        cookies=cookies,
        headers=headers,
    )
    return response

def check_data_existance(file_name):
    return file_name in data_files


while cur_date <= end_date:

    print(cur_date.date(), end = " ") 

    if cur_date.weekday() < 5:
        response = get_response(cur_date)
        data = eval(response.content[15:-1:]) # get data from json callback 
        df = pd.DataFrame(data[1]["analytics"])
        file_name = f"data_{cur_date.date()}.csv"

        if not check_data_existance(file_name):
            df.to_csv(f'Data/index_data/daily_data/{file_name}', index=False)

            print(response.status_code) 
        
        else:
            print("EXISTS")

    else: 
        print("NO DATA") 

    cur_date += timedelta(days=1)

