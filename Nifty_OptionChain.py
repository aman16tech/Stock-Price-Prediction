import time

import gspread
import pandas as pd
import requests
import xlwings as xw
from datetime import datetime

sa = gspread.service_account(r"C:\Users\User\PycharmProjects\pythonProject\Keys.json")
sh = sa.open("Options Data")
wks = sh.worksheet("Sheet1")

while True:
    #connecting to NSE site
    url = 'https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY'
    headers = {
        'user-agent': 'Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.102 Safari/537.36',
        'accept-language': 'en-US,en;q=0.9,bn;q=0.8',
        'accept-encoding': 'gzip, deflate, br'}
    #Requesting for session
    session = requests.session()
    request = session.get(url,headers=headers).json()
    #print(data)
    #cookies= dict(request.cookies)
    #response = session.get(url,headers=headers,cookies=cookies).json()
    rawdata = pd.DataFrame(request)
    #print(rawdata)
    #Got raw dataframework

    rawop = pd.DataFrame(rawdata['filtered']['data']).fillna('')
    #print(rawop)

    strike_price = pd.DataFrame(rawop['strikePrice'])




    def dataframe(rawop):
        data=[]
        for i in range(0,len(rawop)):
            calloi = callcoi = cltp = putoi = putcoi = pltp = callvlm = putvlm =  0
            stp = rawop['strikePrice'][i]
            if(rawop['CE'][i]==0):
                calloi = callcoi=0
            else:
                calloi =rawop['CE'][i]['openInterest']
                callcoi =rawop['CE'][i]['changeinOpenInterest']
                cltp = rawop['CE'][i]['lastPrice']
                callvlm = rawop['CE'][i]['totalTradedVolume']
            if (rawop['PE'][i] == 0):
                putoi = putcoi = 0
            else:
                putoi = rawop['PE'][i]['openInterest']
                putcoi = rawop['PE'][i]['changeinOpenInterest']
                pltp = rawop['PE'][i]['lastPrice']
                putvlm = rawop['PE'][i]['totalTradedVolume']


                PCR = (putoi / calloi) if calloi != 0 else 0

                opdata={'CALL OI':calloi,'CALL COI':callcoi,'CALL LTP':cltp,'PUT OI':putoi,'PUT COI':putcoi,'PUT LTP':pltp,'PC Ratio':PCR, 'CE Volume':callvlm,'PE Volume':putvlm }
                data.append(opdata)
        optionchain = pd.DataFrame(data)
        return  optionchain

    optionchain = dataframe(rawop)

    final = pd.concat([strike_price,optionchain],axis=1)


    #print(final)
    #final.to_csv(r"C:\Users\User\PycharmProjects\pythonProject\Live_OptionsChain_Data.csv")
    #final.to_csv("Live_OptionsChain_Data.csv")
    #time.sleep(5)
    #wb = xw.Book("Live_OptionsChain_Data.csv").sheets['Sheet1']
    #wb.range('A1').expand().value = final

    lst = final.values.tolist()
    wks.update('B2', lst)
    time.sleep(10)
