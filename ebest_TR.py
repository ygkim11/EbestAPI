import win32com.client
import pythoncom
import pandas as pd
import time

class MyObjects:
    server = "demo" # hts : 실서버, demo: 모의
    credentials = pd.read_csv("./credentials/credentials.csv", index_col=0, dtype=str).loc[server, :]

    #for waiting
    login_ok = False
    tr_ok = False  # < TR요청

    stock_code_list = []  # < 종목코드 모아놓는 리스트
    stock_futures_code_list = [] # 주식선물 코드
    stock_futures_basecode_list = [] # 주식선물 기초자산 종목코드
    trade_data = []

    trade_update_cnt = 0
    threshold = None

    ####### 요청 함수 모음
    tr_event = None  # < TR요청에 대한 API 정보

    t8412_request = None  # < 차트데이터 조회 요청함수
    ##################

# TR
class XQ_event_handler:
    def OnReceiveData(self, code):
        print("%s 수신" % code, flush=True)

        if code == "t8436":  # 전종목 종목코드 "005930"
            occurs_count = self.GetBlockCount("t8436OutBlock")

            for i in range(occurs_count):
                shcode = self.GetFieldData("t8436OutBlock", "shcode", i)
                MyObjects.stock_code_list.append(shcode)

            print("주식 종목 리스트: %s" % MyObjects.stock_code_list, flush=True)
            print("주식 종목 갯수: %s" % occurs_count, flush=True)
            MyObjects.tr_ok = True

        elif code == "t8401":  # 주식선물 종목코드
            occurs_count = self.GetBlockCount("t8401OutBlock")

            for i in range(occurs_count):
                shcode = self.GetFieldData("t8401OutBlock", "shcode", i)
                basecode = self.GetFieldData("t8401OutBlock", "basecode", i)
                MyObjects.stock_futures_code_list.append(shcode)
                MyObjects.stock_futures_basecode_list.append(basecode)

            ### !!!추후 최근월물/ 차근월물만 뽑아내는 종목리스트로 바꾸기??? ###

            print("주식선물 종목 리스트: %s" % MyObjects.stock_futures_code_list, flush=True)
            print("주식선물 basecode: %s" % MyObjects.stock_futures_basecode_list, flush=True)
            print("주식선물 종목 갯수: %s" % occurs_count, flush=True)
            MyObjects.tr_ok = True

        elif code == "t8412": # 주식차트 (N분)

            shcode = self.GetFieldData("t8412OutBlock", "shcode", 0)  # 단축코드
            cts_date = self.GetFieldData("t8412OutBlock", "cts_date", 0)  # 연속일자
            cts_time = self.GetFieldData("t8412OutBlock", "cts_time", 0)  # 연속시간

            occurs_count = self.GetBlockCount("t8412OutBlock1")
            for i in range(occurs_count):
                date = self.GetFieldData("t8412OutBlock1", "date", i)
                time = self.GetFieldData("t8412OutBlock1", "time", i)
                close = self.GetFieldData("t8412OutBlock1", "close", i)

                MyObjects.trade_data.append([date, time, close])

            print(shcode, len(MyObjects.trade_data))
            MyObjects.trade_update_cnt += 1


            # 과거 데이터를 더 가져오고 싶을 때는 연속조회를 해야한다.
            if self.IsNext is True:  # 과거 데이터가 더 존재한다.
                if MyObjects.threshold is None or MyObjects.trade_update_cnt < MyObjects.threshold:
                    print("연속조회 기준 날짜: %s" % cts_date, flush=True)
                    MyObjects.t8412_request(shcode=shcode, cts_date=cts_date, cts_time=cts_time, next=self.IsNext)

                # elif MyObjects.trade_update_cnt >= MyObjects.threshold:
                #     MyObjects.tr_ok = True
                else:
                    MyObjects.tr_ok = True
            elif self.IsNext is False:
                MyObjects.tr_ok = True

    def OnReceiveMessage(self, systemError, messageCode, message):
        print("systemError: %s, messageCode: %s, message: %s" % (systemError, messageCode, message))

# Server connection and Login
class XS_event_handler:
    def OnLogin(self, szCode, szMsg):
        print("%s %s" % (szCode, szMsg), flush=True)
        if szCode == "0000":
            MyObjects.login_ok = True
        else:
            MyObjects.login_ok =False

class TR_Main:
    def __init__(self, tr_code, universe="all", threshold=None):
        print("Initiating Main Class")
        print(MyObjects.server , "### 모의투자 ####" if MyObjects.server == "demo" else "@@@@ 실계좌 @@@@")

        #요청 횟수 설정: None 이면 all
        MyObjects.threshold = threshold

        #universe setting
        universe_dict = {}
        universe_dict["all"] = "0"
        universe_dict["KOSPI"] = "1"
        universe_dict["KOSDAQ"] = "2"


        #Login Session
        session = win32com.client.DispatchWithEvents("XA_Session.XASession", XS_event_handler)
        session.ConnectServer(MyObjects.server + ".ebestsec.co.kr", 20001)
        session.Login(MyObjects.credentials["ID"], MyObjects.credentials["PW"], MyObjects.credentials["gonin_PW"], 0, False)

        while MyObjects.login_ok is False:
            pythoncom.PumpWaitingMessages()

        #TR: 주식 종목코드 가져오기 (all, KOSPI, KOSDAQ)
        MyObjects.tr_event = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XQ_event_handler)

        MyObjects.tr_event.ResFileName = "C:/eBEST/xingAPI/Res/t8436.res"
        MyObjects.tr_event.SetFieldData("t8436InBlock", "gubun", 0, universe_dict[universe])
        MyObjects.tr_event.Request(False)

        MyObjects.tr_ok = False
        while MyObjects.tr_ok is False:
            pythoncom.PumpWaitingMessages()

        # TR: 주식선물 종목코드 가져오기
        MyObjects.tr_event = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XQ_event_handler)

        MyObjects.tr_event.ResFileName = "C:/eBEST/xingAPI/Res/t8401.res"
        MyObjects.tr_event.SetFieldData("t8401InBlock", "dummy", 0, "")
        MyObjects.tr_event.Request(False)

        MyObjects.tr_ok = False
        while MyObjects.tr_ok is False:
            pythoncom.PumpWaitingMessages()


        # t8412: 주식차트 N분 요청
        if tr_code == 't8412':
            MyObjects.tr_event.ResFileName = "C:/eBEST/xingAPI/Res/t8412.res"
            MyObjects.t8412_request = self.t8412_request
            for code in MyObjects.stock_code_list:
                MyObjects.trade_data = []
                MyObjects.trade_update_cnt = 0
                MyObjects.t8412_request(shcode=code, cts_date="", cts_time="", next=False)

        # t8412: 주식선물 틱분별 체결조회
        elif tr_code == 't8406':
            pass

        else:
            print("give appropriate tr_code")


    def t8412_request(self, shcode=None, cts_date=None, cts_time=None, next=None):

        time.sleep(1.1)

        MyObjects.tr_event.SetFieldData("t8412InBlock", "shcode", 0, shcode)
        MyObjects.tr_event.SetFieldData("t8412InBlock", "ncnt", 0, 1)
        MyObjects.tr_event.SetFieldData("t8412InBlock", "qrycnt", 0, 500)
        MyObjects.tr_event.SetFieldData("t8412InBlock", "nday", 0, "0")
        MyObjects.tr_event.SetFieldData("t8412InBlock", "sdate", 0, "")
        MyObjects.tr_event.SetFieldData("t8412InBlock", "edate", 0, "당일")
        MyObjects.tr_event.SetFieldData("t8412InBlock", "cts_date", 0, cts_date)
        MyObjects.tr_event.SetFieldData("t8412InBlock", "cts_time", 0, cts_time)
        MyObjects.tr_event.SetFieldData("t8412InBlock", "comp_yn", 0, "N")

        MyObjects.tr_event.Request(next)

        MyObjects.tr_ok = False
        while MyObjects.tr_ok is False:
            pythoncom.PumpWaitingMessages()


if __name__ == "__main__":
    TR_Main(tr_code="t8412", universe="KOSPI", threshold=2)
