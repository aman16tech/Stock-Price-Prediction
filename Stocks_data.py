import gspread
import pandas as pd
import requests,time
from datetime import datetime

sa = gspread.service_account(r"C:\Users\User\PycharmProjects\pythonProject\Keys.json")
sh = sa.open("Live Stock Data")
wks = sh.worksheet("Sheet1")

while True:

    def datetotimestamp(date):
        time_tuple = date.timetuple()
        timestamp = round(time.mktime(time_tuple))
        return timestamp


    def timestamptodate(timestamp):
        return datetime.fromtimestamp(timestamp)


    start = datetotimestamp(datetime(2022, 1, 31))

    end = datetotimestamp(datetime.today())

    url = 'https://priceapi.moneycontrol.com/techCharts/techChartController/history?symbol=ZOMATO&resolution=1&from=' + str(
        start) + '&to=' + str(end) + ''

    resp = requests.get(url).json()
    data = pd.DataFrame(resp)

    date = []
    for dt in data['t']:
        date.append({'Date': timestamptodate(dt)})

    dt = pd.DataFrame(date).astype(str)

    intraday_data = pd.concat([dt, data['o'], data['h'], data['l'], data['c'], data['v']], axis=1).rename(
        columns={'o': 'Open', 'h': 'High', 'l': 'Low', 'c': 'Close', 'v': 'Volume'})
    #print(intraday_data)
    df1 = pd.DataFrame(intraday_data)

    #df1.to_excel("Live_Record.xlsx")
    lst = df1.values.tolist()
    #print(lst)

    wks.update('B4',lst)
    #time.sleep(5)
