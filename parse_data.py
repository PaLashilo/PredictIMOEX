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
    '_ym_isad': '2',
    '_gid': 'GA1.2.604424102.1704216737',
    '_ga': 'GA1.1.41125269.1698584361',
    '_ga_9HDREVDWE7': 'GS1.1.1704221073.13.1.1704221081.52.0.0',
    'MicexPassportCert': 'wTPAAw6TTv6QRx01z38xYwIAAABY9zEbBbTi4zpFPcEiuiyT0ZyJMtJO2tmYCOWxUPwSbP5oVYJf9SYHaOAcoQ9LYg1WNIbIGnK7ETt_M5hSDmyLJAxEw8AOv6qI-VFe_sYaDeFRxwPbgH4iYFo67xLZPLdF5Sa8xSuWQlCx2oGiWuhGJeOn3SZZcSy13mVS6bG_m6oQqYO60',
}

headers = {
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'ru,en;q=0.9,hu;q=0.8,es;q=0.7',
    'Cache-Control': 'no-cache',
    'Connection': 'keep-alive',
    # 'Cookie': '_ym_uid=1682513184127487068; _ym_d=1698584361; _ga_1RFNGQ1758=GS1.1.1702731461.9.0.1702731479.42.0.0; tmr_lvid=7a3ea67c015dc781a31a735757421fc3; tmr_lvidTS=1703082671793; MPTrack=2ea106bf.60cf33d6bd2d2; MicexTrackID=90f7fa46.60d061559ee07; _ym_isad=2; _gid=GA1.2.604424102.1704216737; _ga=GA1.1.41125269.1698584361; _ga_9HDREVDWE7=GS1.1.1704221073.13.1.1704221081.52.0.0; MicexPassportCert=wTPAAw6TTv6QRx01z38xYwIAAABY9zEbBbTi4zpFPcEiuiyT0ZyJMtJO2tmYCOWxUPwSbP5oVYJf9SYHaOAcoQ9LYg1WNIbIGnK7ETt_M5hSDmyLJAxEw8AOv6qI-VFe_sYaDeFRxwPbgH4iYFo67xLZPLdF5Sa8xSuWQlCx2oGiWuhGJeOn3SZZcSy13mVS6bG_m6oQqYO60',
    'Origin': 'https://www.moex.com',
    'Pragma': 'no-cache',
    'Referer': 'https://www.moex.com/',
    'Sec-Fetch-Dest': 'empty',
    'Sec-Fetch-Mode': 'cors',
    'Sec-Fetch-Site': 'same-site',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.5993.732 YaBrowser/23.11.1.732 Yowser/2.5 Safari/537.36',
    'sec-ch-ua': '"Chromium";v="118", "YaBrowser";v="23", "Not=A?Brand";v="99"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
}

params = [
    ('iss.meta', 'off'),
    ('iss.json', 'extended'),
    ('start', '0'),
    ('limit', '100'),
    ('lang', 'ru'),
    # ('date', '2013-11-07'),
    ('nocache-guid', 'f4cdd449-17ea-4bb9-9d68-a95abc1bf3ed'),
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
        if len(df) > 0:
            file_name = f"data_{cur_date.date()}.csv"

            if not check_data_existance(file_name):
                df.to_csv(f'Data/index_data/daily_data/{file_name}', index=False)

                print(response.status_code) 
            
            else:
                print("EXISTS")
        else:
            print("EMPTY RESPONSE")

    else: 
        print("NO DATA") 

    cur_date += timedelta(days=1)

