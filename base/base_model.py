import win32com.client
import os
import time
from pywinauto import application

class BaseModel():
    def __init__(self, id = 'ID', pwd = 'PW', pwdcert = 'PWDCERT'):
        # log in CREON
        self.id = id
        self.pwd = pwd
        self.pwdcert = pwdcert

    def login(self,):
        os.system('taskkill /IM coStarter* /F /T')
        os.system('taskkill /IM CpStart* /F /T')
        os.system('taskkill /IM DibServer* /F /T')
        os.system('wmic process where "name like \'%coStarter%\'" call terminate')
        os.system('wmic process where "name like \'%CpStart%\'" call terminate')
        os.system('wmic process where "name like \'%DibServer%\'" call terminate')
        time.sleep(5)

        app = application.Application()
        app.start(f'C:\CREON\STARTER\coStarter.exe /prj:cp /id:{self.id} /pwd:{self.pwd} /pwdcert:{self.pwdcert} /autostart')
        time.sleep(40)
        print("WELCOME TO CREON")

        # dispatch 
        self.g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
        self.g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
        if self.g_objCpStatus.IsConnect == 0:
            raise Exception("Creon API Connection Error")