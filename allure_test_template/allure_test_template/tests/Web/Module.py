import time

from .Elements import *

class Module:

    def GetTimeStr(self):
        time.sleep(1)
        now = time.strftime(r"%Y-%m-%d")
        return now

    def Login_page(self, Account, Password):
        WaitSendEle('userName', keyList=[Account])
        WaitSendEle('Password', keyList=[Password])
        WaitClickEle('Login')
        time.sleep(10)

# Module()
# ====================================================================================================
mod = Module()
GetTimeStr = mod.GetTimeStr
Login_page = mod.Login_page
