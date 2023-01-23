from sqlalchemy import create_engine
import win32com.client
import pandas as pd
import time
from multiprocessing import Pool
import datetime as dt
from base import BaseModel

class DataLoader(BaseModel):
    def __init__(self,):
        super().__init__()

    def request_code(self,):
        # request the code of listed companies
        kp = self.g_objCodeMgr.GetGroupCodeList(180)  # kospi200
        kq = self.g_objCodeMgr.GetGroupCodeList(390)  # kosdaq150
        return kp, kq

    def request_stock(self, code:str, dvm:str, size:int, count):
        objStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
        objStockChart.SetInputValue(0, code)  # code
        objStockChart.SetInputValue(1, ord('1'))  # 1: period, 2: counts
        objStockChart.SetInputValue(2, 20230101)
        objStockChart.SetInputValue(3, dt.datetime.today().strftime("%Y%m%d"))
        objStockChart.SetInputValue(4, count)  # counts
        objStockChart.SetInputValue(5, [0, 1, 5, 8, 9, 10, 11]) # date, time, close, volume, tran, buy, sell
        objStockChart.SetInputValue(6, ord(str(dvm)))  # dvm: tick
        objStockChart.SetInputValue(7, size)  # size of bar
        objStockChart.SetInputValue(9, ord('1'))  # 0: default, 1: adjusted

        stacked = 0
        date = []; time = []; price = []; vol = []; tran = []; buy = []; sell = []

        while True:
            objStockChart.BlockRequest()
            samp = objStockChart.GetHeaderValue(3)
            if samp == 0:
                break
            for i in range(samp):
                date.append(objStockChart.GetDataValue(0, i))
                time.append(objStockChart.GetDataValue(1, i))
                price.append(objStockChart.GetDataValue(2, i))
                vol.append(objStockChart.GetDataValue(3, i))
                tran.append(objStockChart.GetDataValue(4, i))
                buy.append(objStockChart.GetDataValue(5, i))
                sell.append(objStockChart.GetDataValue(6, i))
            stacked += samp
            print(f'total {stacked} ticks been requested')
            #self._wait()

        df = self._transform_df([date, time, price, vol, tran, buy, sell], code)
        self._store_data(df, 0)

    def request_future(self, code: str, dvm: str, size: int, count):
        objFutureChart = win32com.client.Dispatch("CpSysDib.FutOptChart")
        objFutureChart.SetInputValue(0, code)  # code
        objFutureChart.SetInputValue(1, ord('1'))  # 1: period, 2: counts
        objFutureChart.SetInputValue(2, 20230101)
        objFutureChart.SetInputValue(3, dt.datetime.today().strftime("%Y%m%d"))
        objFutureChart.SetInputValue(4, count)  # counts
        objFutureChart.SetInputValue(5, [0, 1, 5, 8, 9, 10, 11]) # date, time, close, volume, tran, buy, sell
        objFutureChart.SetInputValue(6, ord(str(dvm)))  # dvm: tick
        objFutureChart.SetInputValue(7, size)  # size of bar
        objFutureChart.SetInputValue(8, ord('0'))  # 갭보정
        objFutureChart.SetInputValue(9, ord('1'))  # 0: default, 1: adjusted

        stacked = 0
        date = []; time = []; price = []; vol = []; tran = []; buy = []; sell = []

        while True:
            objFutureChart.BlockRequest()
            samp = objFutureChart.GetHeaderValue(3)
            if samp == 0:
                break
            for i in range(samp):
                date.append(objFutureChart.GetDataValue(0, i))
                time.append(objFutureChart.GetDataValue(1, i))
                price.append(objFutureChart.GetDataValue(2, i))
                vol.append(objFutureChart.GetDataValue(3, i))
                tran.append(objFutureChart.GetDataValue(4, i))
                buy.append(objFutureChart.GetDataValue(5, i))
                sell.append(objFutureChart.GetDataValue(6, i))
            stacked += samp
            #self._wait()
            print(f'{stacked} ticks been requested')
        df = self._transform_df([date, time, price, vol, tran, buy, sell], code)
        self._store_data(df, 1)

    def _transform_df(self, data, code):
        df = pd.DataFrame(data).transpose()[::-1].reset_index(drop=True)
        df.columns = ['date', 'time', 'price', 'vol', 'tran', 'buy', 'sell']
        df.loc[:,:'time'] = df.loc[:,:'time'].astype(int)
        df.loc[:,'price':'sell'] = df.loc[:,'price':'sell'].astype(float)
        df.code = code
        return df

    def _store_data(self, data, type):  # type 0: stock, 1: index_future
        if type == 0:
            name = self.g_objCodeMgr.CodeToName(data.code)
            db = 'mysql+pymysql://root:Sanghunkim25!@127.0.0.1:3306/stocks'
            db_conn = create_engine(db, encoding='utf-8')
            conn = db_conn.connect()
            data.to_sql(name=name, con=db_conn, if_exists='append', index=False)
            conn.close()
            print("Saved the stock sussessfully!")
        else:
            db = 'mysql+pymysql://root:Sanghunkim25!@127.0.0.1:3306/futures'
            db_conn = create_engine(db, encoding='utf-8')
            conn = db_conn.connect()
            data.to_sql(name='future', con=db_conn, if_exists='append', index=False)
            conn.close()
            print("Saved the future sussessfully!")

    def _wait(self,):
        time_remained = self.g_objCpStatus.LimitRequestRemainTime
        cnt_remained = self.g_objCpStatus.GetLimitRemainCount(1)  # 0: order, 1: data_req, 2: live_req
        if cnt_remained <= 0:
            while cnt_remained <= 0:
                time.sleep(time_remained / 1000)
                time_remained = self.g_objCpStatus.LimitRequestRemainTime
                cnt_remained = self.g_objCpStatus.GetLimitRemainCount(1)


if __name__ == "__main__":
    Agent = DataLoader()
    Agent.login()

    # save future ticks
    Agent.request_future("10100", "T", 1, 50000000)

    # save stock ticks
    kp, kq = Agent.request_code()
    for x in kp+kq:
        Agent.request_stock(x, "T", 1, 50000000)