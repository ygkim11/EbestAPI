import win32com.client
import pythoncom

class MyObjects:
    server = "hts" # hts : 실서버, demo: 모의
    login_ok = False

# Real
class XR_event_handler:
    pass
# TR
class XQ_event_handler:
    pass
# Server connection and Login
class XS_event_handler:
    def OnLogin(self, szCode, szMsg):
        print("%s %s" % (szCode, szMsg), flush=True)
        if szCode == "0000":
            MyObjects.login_ok = True
        else:
            MyObjects.login_ok =False

class Main:
    def __init__(self):
        print("Initiating Main Class")
        session = win32com.client.DispatchWithEvents("XA_Session.XASession", XS_event_handler)
        session.ConnectServer(MyObjects.server + ".ebestsec.co.kr", 20001)
        session.Login("jdh00055", "chun6812", "공인", 0, False)

        while MyObjects.login_ok is False:
            pythoncom.PumpWaitingMessages()

if __name__ == "__main__":
    Main()
