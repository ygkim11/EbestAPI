import win32com.client
import pythoncom
import time
import threading
import pandas as pd
import numpy as np

class MyObjects:
    server = "demo"  # hts:실투자, demo: 모의투자
    credentials = pd.read_csv("./credentials/credentials.csv", index_col=0, dtype=str).loc[server, :]

    login_ok =False # Login
    tr_ok = False  # TR요청
    real_ok = False  # 실시간 요청
    acc_no_stock = credentials["acc_no_stocks"]  # 주식 계좌번호
    acc_no_future = credentials["acc_no_futures"]  # 주식선물 계좌번호
    acc_pw = credentials["acc_pw"]  # 계좌비밀번호

    code_list = []
    stock_total_code_list = []  # < 종목코드 모아놓는 리스트
    stock_KOSPI_code_list = []
    stock_KOSDAQ_code_list = []
    stock_futures_code_list = []  # 주식선물 코드
    stock_futures_basecode_list = []  # 주식선물 기초자산 종목코드
    stock_futures_basecode_dict = {}

    stock_trade_order_dict = {}
    futures_trade_order_dict = {}

    # Make Array description
    # ["datetime", "current_price", "open", "high", "low", "volume","cum_volume", "trade_sell_hoga1", "trade_buy_hoga1", "hogatime", "buy_hoga1~10", "sell_hoga1~10", "buy_hoga_stack1~10", "sell_hoga_stack1~10"]
    # [ 0       , 1                 ,2      ,3      ,4      ,5     , 6            , 7                 ,8                  ,9        ,10~19           , 20~29         , 30 ~ 39             , 40~ 49 ]

    #### 요청 함수 모음
    tr_event = None  # TR요청에 대한 API 정보
    real_event_KOSPI = None  # 실시간 요청에 대한 API 정보
    real_event__KOSPI_hoga = None  # 실시간 요청에 대한 API 정보
    real_event_KOSDAQ = None  # 실시간 요청에 대한 API 정보
    real_event_KOSDAQ_hoga = None  # 실시간 요청에 대한 API 정보
    real_event_fu = None
    real_event_fu_hoga = None

    t8412_request = None  # 차트데이터 조회 요청함수
    ##################

# 실시간으로 수신받는 데이터를 다루는 구간
class XR_event_handler:

    def OnReceiveRealData(self, code):

        if code == "K3_": # KOSDAQ 주식체결

            shcode = self.GetFieldData("OutBlock", "shcode")

            if shcode not in MyObjects.stock_trade_order_dict.keys():
                return print("주식 체결: 종목리스트에 없는 shcode 수신...")

            tt = MyObjects.stock_trade_order_dict[shcode]
            tt[0] = self.GetFieldData("OutBlock", "chetime")
            tt[1] = int(self.GetFieldData("OutBlock", "price"))
            tt[2] = int(self.GetFieldData("OutBlock", "open"))
            tt[3] = int(self.GetFieldData("OutBlock", "high"))
            tt[4] = int(self.GetFieldData("OutBlock", "low"))
            tt[5] = int(self.GetFieldData("OutBlock", "cvolume"))
            tt[6] = int(self.GetFieldData("OutBlock", "volume"))
            tt[7] = int(self.GetFieldData("OutBlock", "offerho"))
            tt[8] = int(self.GetFieldData("OutBlock", "bidho"))

            # print(shcode, MyObjects.stock_trade_order_dict[shcode])
            start_time = time.time()
            make_array = np.array(list(MyObjects.stock_trade_order_dict.values()))
            end_time = time.time()

            print(str(start_time - end_time), make_array)


        elif code == "HA_": # KOSDAQ 주식 호가

            shcode = self.GetFieldData("OutBlock", "shcode")
            if shcode not in MyObjects.stock_trade_order_dict.keys():
                return print("주식 호가: 종목리스트에 없는 shcode 수신...")

            tt = MyObjects.stock_trade_order_dict[shcode]
            tt[9] = int(self.GetFieldData("OutBlock", "hotime"))
            for i in range(10,20):
                tt[i] = int(self.GetFieldData("OutBlock", "bidho" + str(i-9)))
                tt[i+20] = int(self.GetFieldData("OutBlock", "bidrem"+str(i-9)))
            for i in range(20,30):
                tt[i] = int(self.GetFieldData("OutBlock", "offerho" + str(i-19)))
                tt[i+20] = int(self.GetFieldData("OutBlock", "offerrem"+str(i-19)))

            # print(shcode, MyObjects.stock_trade_order_dict[shcode])
            start_time = time.time()
            make_array = np.array(list(MyObjects.stock_trade_order_dict.values()))
            end_time = time.time()

            print(str(start_time - end_time), make_array)

        elif code == "S3_": # KOSPI 주식체결

            shcode = self.GetFieldData("OutBlock", "shcode")

            if shcode not in MyObjects.stock_trade_order_dict.keys():
                return print("주식 체결: 종목리스트에 없는 shcode 수신...")

            tt = MyObjects.stock_trade_order_dict[shcode]
            tt[0] = self.GetFieldData("OutBlock", "chetime")
            tt[1] = int(self.GetFieldData("OutBlock", "price"))
            tt[2] = int(self.GetFieldData("OutBlock", "open"))
            tt[3] = int(self.GetFieldData("OutBlock", "high"))
            tt[4] = int(self.GetFieldData("OutBlock", "low"))
            tt[5] = int(self.GetFieldData("OutBlock", "cvolume"))
            tt[6] = int(self.GetFieldData("OutBlock", "volume"))
            tt[7] = int(self.GetFieldData("OutBlock", "offerho"))
            tt[8] = int(self.GetFieldData("OutBlock", "bidho"))

            # print(shcode, MyObjects.stock_trade_order_dict[shcode])

            start_time = time.time()
            make_array = np.array(list(MyObjects.stock_trade_order_dict.values()))
            end_time = time.time()

            print(str(start_time-end_time), make_array)

        elif code == "H1_": # KOSPI 주식 호가

            shcode = self.GetFieldData("OutBlock", "shcode")

            if shcode not in MyObjects.stock_trade_order_dict.keys():
                return print("주식 호가: 종목리스트에 없는 shcode 수신...")

            tt = MyObjects.stock_trade_order_dict[shcode]
            tt[9] = int(self.GetFieldData("OutBlock", "hotime"))
            for i in range(10, 20):
                tt[i] = int(self.GetFieldData("OutBlock", "bidho" + str(i-9)))
                tt[i + 20] = int(self.GetFieldData("OutBlock", "bidrem" + str(i-9)))
            for i in range(20, 30):
                tt[i] = int(self.GetFieldData("OutBlock", "offerho" + str(i-19)))
                tt[i + 20] = int(self.GetFieldData("OutBlock", "offerrem" + str(i-19)))

            # print(shcode, MyObjects.stock_trade_order_dict[shcode])
            start_time = time.time()
            make_array = np.array(list(MyObjects.stock_trade_order_dict.values()))
            end_time = time.time()

            print(str(start_time - end_time), make_array)

        elif code == "JC0": # 주식선물 체결

            futcode = self.GetFieldData("OutBlock", "futcode")

            if futcode not in MyObjects.futures_trade_order_dict.keys():
                return print("주식선물 체결: 종목리스트에 없는 futcode 수신...")

            tt = MyObjects.futures_trade_order_dict[futcode]
            tt[0] = self.GetFieldData("OutBlock", "chetime")
            tt[1] = int(self.GetFieldData("OutBlock", "price"))
            tt[2] = int(self.GetFieldData("OutBlock", "open"))
            tt[3] = int(self.GetFieldData("OutBlock", "high"))
            tt[4] = int(self.GetFieldData("OutBlock", "low"))
            tt[5] = int(self.GetFieldData("OutBlock", "cvolume"))
            tt[6] = int(self.GetFieldData("OutBlock", "volume"))
            tt[7] = int(self.GetFieldData("OutBlock", "offerho1"))
            tt[8] = int(self.GetFieldData("OutBlock", "bidho1"))

            # print(futcode, MyObjects.futures_trade_order_dict[futcode])
            start_time = time.time()
            make_array = np.array(list(MyObjects.futures_trade_order_dict.values()))
            end_time = time.time()

            print(str(start_time - end_time), make_array)

        elif code == "JH0": # 주식선물 호가

            futcode = self.GetFieldData("OutBlock", "futcode")

            if futcode not in MyObjects.futures_trade_order_dict.keys():
                return print("주식선물 호가: 종목리스트에 없는 futcode 수신...")

            tt = MyObjects.futures_trade_order_dict[futcode]
            tt[9] = int(self.GetFieldData("OutBlock", "hotime"))
            for i in range(10, 20):
                tt[i] = int(self.GetFieldData("OutBlock", "bidho" + str(i-9)))
                tt[i + 20] = int(self.GetFieldData("OutBlock", "bidrem" + str(i-9)))
            for i in range(20, 30):
                tt[i] = int(self.GetFieldData("OutBlock", "offerho" + str(i-19)))
                tt[i + 20] = int(self.GetFieldData("OutBlock", "offerrem" + str(i-19)))

            # print(futcode, MyObjects.futures_trade_order_dict[futcode])
            start_time = time.time()
            make_array = np.array(list(MyObjects.futures_trade_order_dict.values()))
            end_time = time.time()

            print(str(start_time - end_time), make_array)

# TR 요청 이후 수신결과 데이터를 다루는 구간
class XQ_event_handler:

    def OnReceiveData(self, code):
        print("%s 수신" % code, flush=True)

        if code == "t8436":
            occurs_count = self.GetBlockCount("t8436OutBlock")
            for i in range(occurs_count):
                shcode = self.GetFieldData("t8436OutBlock", "shcode", i)
                MyObjects.code_list.append(shcode)

            print(occurs_count, "주식 종목 리스트: %s" % MyObjects.code_list, flush=True)
            MyObjects.tr_ok = True

        elif code == "t8401":  # 주식선물 종목코드
            occurs_count = self.GetBlockCount("t8401OutBlock")
            # print("주식선물 종목 갯수: %s" % occurs_count, flush=True)

            for i in range(occurs_count):
                shcode = self.GetFieldData("t8401OutBlock", "shcode", i)
                basecode = self.GetFieldData("t8401OutBlock", "basecode", i)
                MyObjects.stock_futures_code_list.append(shcode)

                #make dictionary
                MyObjects.stock_futures_basecode_dict[shcode] = basecode[1:] # basecode 앞의 "A" 제거

            ### 최근월물/ 차근월물만 뽑아내는 종목리스트로 바꾸기 (추후 보완 ㄱ) ###
            fu_code_ls = list(set(map(lambda x: x[1:3], MyObjects.stock_futures_code_list)))

            total_fu_code = []
            for fu_code in fu_code_ls:
                fut_tmp = []
                for i in range(len(MyObjects.stock_futures_code_list)):
                    fu_code_i = MyObjects.stock_futures_code_list[i][1:3]
                    if fu_code_i == fu_code:
                        if MyObjects.stock_futures_code_list[i][0] == "1": # 선물이면 종목코드의 첫자리가 1 이여야 함.
                            fut_tmp.append(MyObjects.stock_futures_code_list[i])
                    else:
                        pass
                total_fu_code.append(fut_tmp)

            total_fu_code = list(map(lambda x: x[:1], total_fu_code))  # 더 원월물까지 포함하고 싶으면 3을 바꾸면됨

            flatten_fu_code = []
            for fu_code in total_fu_code:
                flatten_fu_code = flatten_fu_code + fu_code

            MyObjects.stock_futures_code_list = flatten_fu_code # 주식선물(근월물)만으로 filter된 stock_futures_code_list 생성

            basecode_by_dict = []
            for fu_code in MyObjects.stock_futures_code_list:
                basecode_by_dict.append(MyObjects.stock_futures_basecode_dict[fu_code])

            MyObjects.stock_futures_basecode_list = list(set(basecode_by_dict)) # 주식선물에 대한 base code list

            print(len(MyObjects.stock_futures_code_list), "주식선물 종목 리스트: %s" % MyObjects.stock_futures_code_list, flush=True)
            print(len(MyObjects.stock_futures_basecode_list),"주식선물 basecode: %s" % MyObjects.stock_futures_basecode_list, flush=True)

            MyObjects.tr_ok = True

    def OnReceiveMessage(self, systemError, messageCode, message):
        print("systemError: %s, messageCode: %s, message: %s" % (systemError, messageCode, message), flush=True)

# 서버접속 및 로그인 요청 이후 수신결과 데이터를 다루는 구간
class XS_event_handler:

    def OnLogin(self, szCode, szMsg):
        print("%s %s" % (szCode, szMsg), flush=True)
        if szCode == "0000":
            MyObjects.login_ok = True
        else:
            MyObjects.login_ok = False


# 실행용 클래스
class Main:
    def __init__(self):
        print("실행용 클래스이다")

        session = win32com.client.DispatchWithEvents("XA_Session.XASession", XS_event_handler)
        session.ConnectServer(MyObjects.server + ".ebestsec.co.kr", 20001)  # 서버 연결
        session.Login(MyObjects.credentials["ID"], MyObjects.credentials["PW"], MyObjects.credentials["gonin_PW"], 0, False)  # 서버 연결

        while MyObjects.login_ok is False:
            pythoncom.PumpWaitingMessages()

        # TR: 주식 종목코드 가져오기
        for i in range(0, 3):
            MyObjects.tr_event = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XQ_event_handler)
            MyObjects.tr_event.ResFileName = "C:/eBEST/xingAPI/Res/t8436.res"
            MyObjects.tr_event.SetFieldData("t8436InBlock", "gubun", 0, str(i))
            MyObjects.tr_event.Request(False)

            MyObjects.tr_ok = False
            while MyObjects.tr_ok is False:
                pythoncom.PumpWaitingMessages()

            if i == 0:
                MyObjects.stock_total_code_list = MyObjects.code_list
                time.sleep(1)
            elif i == 1:
                MyObjects.stock_KOSPI_code_list = MyObjects.code_list
                time.sleep(1)
            elif i == 2:
                MyObjects.stock_KOSDAQ_code_list = MyObjects.code_list
                time.sleep(1)

        # TR: 주식선물 종목코드 가져오기
        MyObjects.tr_event = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XQ_event_handler)

        MyObjects.tr_event.ResFileName = "C:/eBEST/xingAPI/Res/t8401.res"
        MyObjects.tr_event.SetFieldData("t8401InBlock", "dummy", 0, "")
        MyObjects.tr_event.Request(False)

        MyObjects.tr_ok = False
        while MyObjects.tr_ok is False:
            pythoncom.PumpWaitingMessages()

        ############For Test################
        MyObjects.stock_code_list = ["005930", "096530"]
        MyObjects.stock_futures_code_list = ["111R2000", "1CLR2000"]
        MyObjects.stock_futures_basecode_list = ["005930", "096530"]
        ####################################

        # KOSDAQ
        MyObjects.real_event_KOSDAQ = win32com.client.DispatchWithEvents("XA_DataSet.XAReal", XR_event_handler)
        MyObjects.real_event_KOSDAQ.ResFileName = "C:/eBEST/xingAPI/Res/K3_.res"
        for shcode in MyObjects.stock_futures_basecode_list:  # 주식선물의 기초자산 종목만 등록!
            if shcode in MyObjects.stock_KOSDAQ_code_list:
                print("KOSDAQ 주식 체결 종목 등록 %s" % shcode)
                MyObjects.stock_trade_order_dict[shcode] = np.zeros(50)
                MyObjects.real_event_KOSDAQ.SetFieldData("InBlock", "shcode", shcode)
                MyObjects.real_event_KOSDAQ.AdviseRealData()

        MyObjects.real_event_KOSDAQ_hoga = win32com.client.DispatchWithEvents("XA_DataSet.XAReal", XR_event_handler)
        MyObjects.real_event_KOSDAQ_hoga.ResFileName = "C:/eBEST/xingAPI/Res/HA_.res"
        for shcode in MyObjects.stock_futures_basecode_list:  # 주식선물의 기초자산 종목만 등록!
            if shcode in MyObjects.stock_KOSDAQ_code_list:
                print("KOSDAQ 주식 호가잔량 종목 등록 %s" % shcode)
                MyObjects.real_event_KOSDAQ_hoga.SetFieldData("InBlock", "shcode", shcode)
                MyObjects.real_event_KOSDAQ_hoga.AdviseRealData()

        # KOSPI
        MyObjects.real_event_KOSPI = win32com.client.DispatchWithEvents("XA_DataSet.XAReal", XR_event_handler)
        MyObjects.real_event_KOSPI.ResFileName = "C:/eBEST/xingAPI/Res/S3_.res"
        for shcode in MyObjects.stock_futures_basecode_list:  # 주식선물의 기초자산 종목만 등록!
            if shcode in MyObjects.stock_KOSPI_code_list:
                print("KOSPI 주식 체결 종목 등록 %s" % shcode)
                MyObjects.stock_trade_order_dict[shcode] = np.zeros(50)
                MyObjects.real_event_KOSPI.SetFieldData("InBlock", "shcode", shcode)
                MyObjects.real_event_KOSPI.AdviseRealData()

        MyObjects.real_event_KOSPI_hoga = win32com.client.DispatchWithEvents("XA_DataSet.XAReal", XR_event_handler)
        MyObjects.real_event_KOSPI_hoga.ResFileName = "C:/eBEST/xingAPI/Res/H1_.res"
        for shcode in MyObjects.stock_futures_basecode_list:  # 주식선물의 기초자산 종목만 등록!
            if shcode in MyObjects.stock_KOSPI_code_list:
                print("KOSPI 주식 호가잔량 종목 등록 %s" % shcode)
                MyObjects.real_event_KOSPI_hoga.SetFieldData("InBlock", "shcode", shcode)
                MyObjects.real_event_KOSPI_hoga.AdviseRealData()

        # 주식선물
        MyObjects.real_event_fu = win32com.client.DispatchWithEvents("XA_DataSet.XAReal", XR_event_handler)
        MyObjects.real_event_fu.ResFileName = "C:/eBEST/xingAPI/Res/JC0.res"
        for futcode in MyObjects.stock_futures_code_list:
            print("주식선물 체결 종목 등록 %s" % futcode)
            MyObjects.futures_trade_order_dict[futcode] = np.zeros(50)
            MyObjects.real_event_fu.SetFieldData("InBlock", "futcode", futcode)
            MyObjects.real_event_fu.AdviseRealData()

        MyObjects.real_event_fu_hoga = win32com.client.DispatchWithEvents("XA_DataSet.XAReal", XR_event_handler)
        MyObjects.real_event_fu_hoga.ResFileName = "C:/eBEST/xingAPI/Res/JH0.res"
        for futcode in MyObjects.stock_futures_code_list:
            print("주식선물 호가잔량 종목 등록 %s" % futcode)
            MyObjects.real_event_fu_hoga.SetFieldData("InBlock", "futcode", futcode)
            MyObjects.real_event_fu_hoga.AdviseRealData()

        while MyObjects.real_ok is False:
            pythoncom.PumpWaitingMessages()

if __name__ == "__main__":
    Main()