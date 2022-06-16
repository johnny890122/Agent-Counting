from email import header
import get_gdoc  # 讀入Google Sheets專用
import time
import numpy as np
import pandas as pd
import datetime
import warnings
import socket
from google.oauth2.service_account import Credentials
import gspread
warnings.filterwarnings('ignore')
socket.setdefaulttimeout(180)  # 超過180秒才會報超時

'''
使用說明 2021/9/29 Mic
1. 將counting_raw.csv放入Input資料夾中
2. 將第19行輸入想要的月份
3. 若進貨異常表單有更動，請更改第38~39行的網址，並注意第40行的範圍是否需更動(第一欄表頭位置可能更動)
'''

time0 = time.time()
month_first_day = datetime.datetime.strptime('2022-05', '%Y-%m')
month = month_first_day.strftime("%b")

scope = ['https://www.googleapis.com/auth/spreadsheets']
creds = Credentials.from_service_account_file("credentials.json", scopes=scope)
gs = gspread.authorize(creds)

class gdoc_information():
    def __init__(self):
        self.SCOPES = []
        self.SAMPLE_SPREADSHEET_ID = []
        self.SAMPLE_RANGE_NAME = []

    def trans(self):
        tmp = []
        tmp.extend(self.SCOPES)
        tmp.extend(self.SAMPLE_SPREADSHEET_ID)
        tmp.extend(self.SAMPLE_RANGE_NAME)
        return tmp


# 異常回覆表單
abs_gdoc = gdoc_information()
abs_gdoc.SCOPES = 'https://docs.google.com/spreadsheets/d/10ylvlT6KzZ9kQi4-VTmx5V7FIVOJs-MaSphuLqoAA5M'
abs_gdoc.SAMPLE_RANGE_NAME = '倉庫回報表格'

abs_raw = gs.open_by_url(abs_gdoc.SCOPES).worksheet(abs_gdoc.SAMPLE_RANGE_NAME).get_all_values()
header_row_idx = 0

while abs_raw[header_row_idx][0] != "Date":
    header_row_idx += 1


abnormal_header = abs_raw[header_row_idx]
abnormal = pd.DataFrame(abs_raw[header_row_idx+1:])
abnormal.columns = abnormal_header

abnormal['Inbound ID'] = abnormal['Inbound ID'].str.upper()
abnormal['Date'] = pd.to_datetime(abnormal['Date'], errors="coerce")
abnormal = abnormal[abnormal['Date'].dt.month == month_first_day.month]

# 取數錯
abnormal_counting = abnormal[(abnormal['組別'].isin(['貼標', '質檢', '驗貨'])) &
                             (abnormal['問題'].isin(['多貨進倉', '數量短少']))]['Inbound ID'].values
# 取沒檢查到包裝
abnormal_packing = abnormal[(abnormal['組別'].isin(['貼標', '質檢', '驗貨'])) &
                            (abnormal['問題'].isin(['商品凹/破', '包裝異常']))]['Inbound ID'].values

counting_raw = pd.read_csv('Input/counting_raw_{}.csv'.format(month)).drop_duplicates(subset=['tracking_id'])
counting_raw['數錯'] = np.where(counting_raw['inbound_id'].isin(abnormal_counting), 1, 0)
counting_raw['沒檢查到包裝'] = np.where(counting_raw['inbound_id'].isin(abnormal_packing), 1, 0)
counting_raw['Operator'] = counting_raw['counting_start_op'].str.replace('@shopee.com', '')

accuracy = counting_raw.groupby('Operator')\
                       .agg({'Operator': 'count', '數錯': np.sum, '沒檢查到包裝': np.sum})\
                       .rename(columns={'Operator': 'Total'})
accuracy['Accuracy'] = 1 - (accuracy['數錯'] + accuracy['沒檢查到包裝']) / accuracy['Total']
accuracy.to_excel('Output/agent_acc_counting_{}.xlsx'.format(month_first_day.strftime("%b")), encoding='utf_8_sig')

time1 = time.time()
print('任務完成，共花費 {:.2f} 秒'.format(time1 - time0))
