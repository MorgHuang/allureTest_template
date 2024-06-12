import os, logging, yaml, time, datetime, allure, json, requests, jwt, re, csv, shutil

import pytest, sys, random
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait as WDW
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains as AC
from selenium.common.exceptions import NoSuchFrameException
from openpyxl import load_workbook


class Elements:

    def __init__(self):
        '''啟動時自動導入'''
        self.LoadData()

    def InIDiver(self, driver):
        '''透過外部參數導入 driver'''
        self.driver = driver

    def LoadData(self):
        '''讀取指定目錄下的 Page.yaml'''
        path = os.path.abspath(os.path.join(os.getcwd(), 'data', 'Web', 'Page.yaml'))
        with open(path, encoding="utf-8") as f:
            try:
                self.elementsData = yaml.safe_load(f)
            except yaml.YAMLError as e:
                logging.error(e)

    def data_yaml(self):
        '''讀取指定目錄下的 data.yaml'''
        yaml_path = os.path.abspath(os.path.join(os.getcwd(), 'data', 'Web', "data.yaml"))
        with open(yaml_path, encoding="utf-8") as f:
            try:
                data = yaml.safe_load(f)
            except yaml.YAMLError as e:
                logging.error(e)
        return yaml_path, data

    def env(self, url):
        '''讀取指定目錄下的 env.yaml'''
        config_path = os.path.abspath(os.path.join(os.getcwd(), "data", "env.yaml"))
        with open(config_path, encoding="utf-8") as f:
            try:
                env = yaml.safe_load(f)[url]
            except yaml.YAMLError as e:
                logging.error(e)
        return env

    def imageDict(self):
        imageDict = os.path.abspath(os.path.join(os.getcwd(), "image"))
        if not os.path.isdir(imageDict):
            os.mkdir(imageDict)
        return imageDict

    # selenium element trigger
    # ====================================================================================================
    def ClickEle(self, eleName, img=True, sleepTime=1):
        '''click'''
        ele, dealError = self.GetElement(eleName)
        assert dealError == True, f'eleName = {eleName}，dealError must be True！'
        if ele is None:  # GetElement = None 時觸發
            self.dealError('ClickEle', eleName, img)
        else:
            try:
                ele.click()
                time.sleep(int(sleepTime))
            except Exception as e:
                self.exc('ClickEle', eleName, e, img)

    def WaitClickEle(self, eleName, img=True):
        '''隱式等待 + click'''
        ele, dealError = self.GetEleExceptionEle(eleName)
        assert dealError == True, f'eleName = {eleName}，dealError must be True！'
        if ele is None:  # GetEleExceptionEle = None 時觸發
            self.dealError('WaitClickEle', eleName, img)
        else:
            try:
                ele.click()
                time.sleep(1)
                logging.info(eleName + ' - WaitClickEle Success')
            except Exception as e:
                self.exc('WaitClickEle', eleName, e, img)

    def Click_Ele_False(self, eleName, img=True, sleep_time=10):
        '''隱式等待 + click + sleep，just for dealError: False'''
        ele, dealError = self.GetEleExceptionEle(eleName, sleep_time)
        assert dealError == False, f'eleName = {eleName}，dealError must be False！'
        if ele is None:  # GetEleExceptionEle = None 時觸發
            logging.error(f'{eleName}，is None')
            return None
        else:
            try:
                ele.click()
                logging.info(f'{eleName} - Click_Ele_False Success')
                time.sleep(1)
                return True
            except Exception as e:
                self.exc('Click_Ele_False', eleName, e, img)

    def GetText(self, eleName, img=True):
        '''GetText'''
        ele, dealError = self.GetElement(eleName)
        assert dealError == True, f'eleName = {eleName}，dealError must be True！'
        if ele is None:  # GetElement = None 時觸發
            self.dealError('GetText', eleName, img)
        else:
            try:
                return ele.text
            except Exception as e:
                self.exc('GetText', eleName, e, img)

    def WaitGetText(self, eleName, img=True):
        '''隱式等待 + GetText'''
        ele, dealError = self.GetEleExceptionEle(eleName)
        assert dealError == True, f'eleName = {eleName}，dealError must be True！'
        if ele is None:  # GetEleExceptionEle = None 時觸發
            self.dealError('WaitGetText', eleName, img)
        else:
            try:  # 嘗試回傳 value
                logging.info('eleName = ' + eleName + '，text = ' + str(ele.text) + ' - WaitGetText Success')
                return str(ele.text)
            except Exception as e:
                self.exc('WaitGetText', eleName, e, img)

    def SendKey_Ele(self, eleName, data, img=True, sleepTime=1):
        '''send_keys'''
        ele, dealError = self.GetElement(eleName)
        assert dealError == True, f'eleName = {eleName}，dealError must be True！'
        if ele is None:  # GetElement = None 時觸發
            self.dealError('SendKey_Ele', eleName, img)
        else:
            try:
                ele.clear()
                ele.send_keys(data)
                time.sleep(sleepTime)
            except Exception as e:
                self.exc('SendKey_Ele', eleName, e, img)

    def WaitSendEle(self, eleName, keyList=[], img=True):
        '''隱式等待 + send_keys'''
        ele, dealError = self.GetEleExceptionEle(eleName)
        assert dealError == True, f'eleName = {eleName}，dealError must be True！'
        if ele is None:  # GetEleExceptionEle = None 時觸發
            self.dealError('WaitSendEle', eleName, img)
        else:
            try:
                # send_keys Success 時觸發
                ele.clear()
                for keyStr in keyList:
                    ele.send_keys(keyStr)
                    time.sleep(1)
                    logging.info('eleName = ' + eleName + '，keyStr = ' + str(keyStr) + ' - WaitSendEle Success')
            except Exception as e:
                self.exc('WaitSendEle', eleName, e, img)

    def SelectEle(self, eleName, by, obj, img=True, sleepTime=1):
        '''
        控制網頁的 Select 標籤
        By: index, value, text
        '''
        ele, dealError = self.GetElement(eleName)
        assert dealError == True, f'eleName = {eleName}，dealError must be True！'
        if ele is None:  # GetElement = None 時觸發
            self.dealError('SelectEle', eleName, img)
        else:
            try:
                if by == "index":
                    Select(ele).select_by_index(obj)
                    time.sleep(int(sleepTime))
                elif by == "value":
                    Select(ele).select_by_value(obj)
                    time.sleep(int(sleepTime))
                elif by == "text":
                    Select(ele).select_by_visible_text(obj)
                    time.sleep(int(sleepTime))
            except Exception as e:
                self.exc('SelectEle', eleName, e, img)

    def Clear_Ele(self, eleName, img=True, sleepTime=1):
        '''清除該物件的內容'''
        ele, dealError = self.GetElement(eleName)
        assert dealError == True, f'eleName = {eleName}，dealError must be True！'
        if ele is None:  # GetElement = None 時觸發
            self.dealError('Clear_Ele', eleName, img)
        else:
            try:
                ele.clear()
                logging.info(eleName + ' - Clear_Ele Success')
                time.sleep(sleepTime)
            except Exception as e:
                self.exc('Clear_Ele', eleName, e, img)

    def Clear_Ele_By_Keys(self, eleName, keyList=[Keys.CONTROL + 'a' + Keys.DELETE], img=True):
        '''隱式等待 + keyList=[Keys.CONTROL+'a'+Keys.DELETE]'''
        ele, dealError = self.GetEleExceptionEle(eleName)
        assert dealError == True, f'eleName = {eleName}，dealError must be True！'
        if ele is None:  # GetEleExceptionEle = None 時觸發
            self.dealError('Clear_Ele_By_Keys', eleName, img)
        else:
            try:
                # send_keys Success 時觸發
                ele.clear()
                for keyStr in keyList:
                    ele.send_keys(keyStr)
                    time.sleep(1)
                    logging.info(f"eleName = {eleName} - Clear_Ele_By_Keys Success")
            except Exception as e:
                self.exc('Clear_Ele_By_Keys', eleName, e, img)

    def exc(self, module, eleName, e, img):
        '''該 Module 沒有正常執行！'''
        if img == True:
            self.Screenshot(f'{module} - {eleName} - except')
        logging.error(f'{module} - {eleName} - {e}，except！')
        raise Exception(f'{module} - {eleName}，except！')

    def dealError(self, module, eleName, img):
        '''dealError == True'''
        if img == True:
            self.Screenshot(f'{module} - {eleName} - dealError')
        logging.error(f'{module} - {eleName}，dealError！')
        raise Exception(f'{module} - {eleName}，dealError！')

    def actions_click(self, eleName):
        ele, _ = self.GetEleExceptionEle(eleName)
        actions = AC(self.driver)
        actions.click(ele)
        actions.perform()
        time.sleep(1)
        logging.info(f'{eleName} - actions_click Success')

    def actions_send_keys(self, eleName, value):
        ele, _ = self.GetEleExceptionEle(eleName)
        actions = AC(self.driver)
        actions.send_keys_to_element(ele, value)
        actions.perform()
        time.sleep(1)
        logging.info(f'eleName = {eleName}，value = {value} - actions_send_keys Success')

    def actions_move(self, eleName):
        ele, _ = self.GetEleExceptionEle(eleName)
        actions = AC(self.driver)
        actions.move_to_element(ele)
        actions.perform()
        time.sleep(1)
        logging.info(f'{eleName} - actions_move Success')

    def actions_move_click(self, eleName):
        ele, _ = self.GetEleExceptionEle(eleName)
        actions = AC(self.driver)
        actions.move_to_element(ele).click()
        actions.perform()
        time.sleep(1)
        logging.info(f'{eleName} - actions_move_click Success')

    def actions_by_offset_click(self, x, y):
        actions = AC(self.driver)
        actions.move_by_offset(x, y).click()
        actions.perform()
        time.sleep(1)
        logging.info(f'{x},{y} - actions_by_offset_click Success')

    def actions_by_offset_click_and_hold(self, eleName, x, y):
        ele, _ = self.GetEleExceptionEle(eleName)
        actions = AC(self.driver)
        actions.click_and_hold(ele)
        actions.move_by_offset(x, y)
        actions.release()
        actions.perform()
        time.sleep(1)
        logging.info(f'{x},{y} - actions_by_offset_click Success')

    # selenium element read
    # ====================================================================================================
    def GetEleExceptionEle(self, ele_name, sleep_time=25):
        '''隱式等待 + 讀取物件設定'''
        try:
            ele_info = self.elementsData[ele_name]
            ele_type = ele_info['type']
            ele_value = ele_info['value']
            deal_error = ele_info.get('dealError', False)
        except KeyError:
            logging.error(f"{ele_name} Incorrect setting!")
            raise Exception(f"{ele_name} Incorrect setting!")

        try:
            locator_dict = {"id": By.ID, "xpath": By.XPATH, "link": By.LINK_TEXT, "partial": By.PARTIAL_LINK_TEXT,
                            "name": By.NAME, "tag": By.TAG_NAME, "class": By.CLASS_NAME, "css": By.CSS_SELECTOR}
            element = WDW(self.driver, sleep_time).until(
                EC.presence_of_element_located((locator_dict[ele_type], ele_value)))
            self.border(element)
            return element, deal_error
        except Exception as e:
            logging.error(f"{ele_name} - {str(e)}")
            return None, deal_error

    def GetElement(self, ele_name):
        '''讀取物件設定'''
        try:
            ele_info = self.elementsData[ele_name]
            ele_type = ele_info['type']
            ele_value = ele_info['value']
            deal_error = ele_info.get('dealError', False)
        except KeyError:
            logging.error(f"{ele_name} Incorrect setting!")
            raise Exception(f"{ele_name} Incorrect setting!")

        try:
            locator_dict = {"id": By.ID, "xpath": By.XPATH, "link": By.LINK_TEXT, "partial": By.PARTIAL_LINK_TEXT,
                            "name": By.NAME, "tag": By.TAG_NAME, "class": By.CLASS_NAME, "css": By.CSS_SELECTOR}
            element = self.driver.find_element(locator_dict[ele_type], ele_value)
            self.border(element)
            return element, deal_error
        except Exception as e:
            logging.error(f"{ele_name} - {str(e)}")
            return None, deal_error

    # function element
    # ====================================================================================================
    def JavaScript(self, js):
        '''透過 driver，讓 browser 執行 JS 指令'''
        self.driver.execute_script(js)
        logging.info('script = ' + '\n' + js + '\n- execute_script Success')

    def JS_Click(self, css):
        self.driver.execute_script("document.querySelector('{}').click()".format(css))
        logging.info('JS_Click = ' + css + ' - JS_Click Success')

    def JS_Value(self, css, value):
        self.driver.execute_script("document.querySelector('{}').value = '{}'".format(css, value))
        logging.info('JS_Value = ' + css + ' - JS_Value Success')

    def SwitchWindow(self, index, sleepTime=2):
        '''所有視窗 ID + 當前視窗 ID + 切換視窗'''
        try:
            handles = self.driver.window_handles
            logging.info('window_handles = ' + str(handles) + '\n')
            logging.info('current_window_handle = ' + str(self.driver.current_window_handle))
            logging.info('title = ' + str(self.driver.title))
            logging.info('current_url = ' + str(self.driver.current_url) + '\n')
            self.driver.switch_to.window(str(handles[index]))
            logging.info('switch_to.window = ' + str(self.driver.current_window_handle))
            logging.info('title = ' + str(self.driver.title))
            logging.info('current_url = ' + str(self.driver.current_url))
            time.sleep(sleepTime)

        except Exception as e:
            logging.error(str(e))

    def SwitchWindow_Url(self, correct_url, sleepTime=2):
        '''
        SwitchWindow 改良版本
        用於比對多個視窗的網址，是否為想要切換的視窗?
        如果是的話終止迴圈，不是的話繼續切換視窗直到找到符合的網址~
        '''
        try:
            handles = self.driver.window_handles
            logging.info('handles = ' + str(handles) + '，len = ' + str(len(handles)) + '\n')
            for handle in handles:
                self.driver.switch_to.window(handle)
                index = list(handles).index(handle)
                time.sleep(sleepTime)
                if self.driver.current_url == correct_url:
                    logging.info('符合')
                    logging.info('handle = ' + str(handle) + '，index = ' + str(index))
                    logging.info('title = ' + str(self.driver.title))
                    logging.info('current_url = ' + str(self.driver.current_url) + '\n')
                    break
                else:
                    logging.info('不符合')
                    logging.info('handle = ' + str(handle) + '，index = ' + str(index))
                    logging.info('title = ' + str(self.driver.title))
                    logging.info('current_url = ' + str(self.driver.current_url) + '\n')
                if handle == handles[-1]: logging.warning('所有網址皆不符合！')

        except Exception as e:
            logging.error(str(e))

    def CloseWindow(self):
        '''關閉當前視窗'''
        self.driver.close()

    def MaximizeWindow(self):
        '''視窗最大化'''
        self.driver.maximize_window()

    def BackWin(self):
        '''上一頁'''
        self.driver.back()

    def ForwardWin(self):
        '''下一頁'''
        self.driver.forward()

    def screenshot(self, filename):
        '''截圖到 local'''
        filenameTime = os.path.join(self.imageDict(), filename + '.png')
        self.driver.get_screenshot_as_file(filenameTime)

    def SwitchToFrame(self, num):
        '''切換框架至下一層(用於透過框架進行定位)'''
        self.driver.switch_to.frame(num)
        logging.info(f'SwitchToFrame = {num} - SwitchToFrame Success')

    def SwitchToParent(self):
        '''切換框架至上一層(用於透過框架進行定位)'''
        self.driver.switch_to.parent_frame()
        logging.info('SwitchToParent Success')

    def SwitchToDefault(self):
        '''切換到預設框架(主框架)'''
        self.driver.switch_to.default_content()
        logging.info('SwitchToDefault Success')

    def Allure(self, filename):
        '''Allure 上傳圖片'''
        filenameTime = os.path.join(self.imageDict(), filename + '.png')
        allure.attach.file(filenameTime, attachment_type=allure.attachment_type.PNG, name=filename)

    def GET_CurUrl(self):
        '''回傳當前網址'''
        url = self.driver.current_url
        logging.info(f'current_url = {url}')
        return url

    def GET_Title(self):
        '''回傳當前 Title'''
        title = self.driver.title
        logging.info('title = ' + title)
        return title

    def GetTimeStr(self):
        '''取得當前日期'''
        time.sleep(1)
        now = time.strftime(r"_%Y-%m-%d")
        return now

    def Enter_Url(self, value):
        '''切換到指定網址'''
        try:
            self.driver.get(value)
            logging.info(value + ' - Enter_Url Success')
            time.sleep(1)
        except Exception as e:
            logging.error(value + ' - ' + str(e))

    def Refresh(self):
        '''重新整理頁面'''
        self.driver.refresh()
        logging.info('Refresh Success')

    def Screenshot(self, name):
        '''screenshot + Allure'''
        image = name + self.GetTimeStr()
        self.screenshot(image)
        self.Allure(image)
        logging.info(image + ' - Screenshot Success')

    def is_displayed(self, eleName):
        '''判斷元素是否顯示'''
        try:
            ele, dealError = self.GetEleExceptionEle(eleName)
            display = ele.is_displayed()
            logging.info(eleName + ' - is_displayed ? ' + str(display))
            return display
        except Exception as e:
            logging.error(eleName + ' - ' + str(e))

    def is_enabled(self, eleName):
        '''判斷元素是否可用'''
        try:
            ele, dealError = self.GetEleExceptionEle(eleName)
            enabled = ele.is_enabled()
            logging.info(eleName + ' - is_enabled ? ' + str(enabled))
            return enabled
        except Exception as e:
            logging.error(eleName + ' - ' + str(e))

    def is_selected(self, eleName):
        '''判斷元素是否被選取'''
        try:
            ele, dealError = self.GetEleExceptionEle(eleName)
            selected = ele.is_selected()
            logging.info(eleName + ' - is_selected ? ' + str(selected))
            return selected
        except Exception as e:
            logging.error(eleName + ' - ' + str(e))

    def GoToWindow(self, url):
        '''開啟當前視窗'''
        try:
            self.driver.get(url)
            logging.info(f'{url}，GoToWindow Success')
        except Exception as e:
            logging.error(str(e))

    def New_openpage(self):
        self.driver.execute_script("window.open('');")

    def switch_to_alert(self):
        '''witch_to.alert > accept'''
        try:
            self.driver.switch_to.alert.accept()
            logging.info('switch_to_alert Success')
        except Exception as e:
            logging.error(str(e))

    def border(self, element):
        '''在物件周圍顯示紅色外框'''
        self.driver.execute_script("arguments[0].style.border = '2px solid red';", element)
        time.sleep(1)
        self.driver.execute_script("arguments[0].style.border = '';", element)


# Elements()
# ====================================================================================================
ele = Elements()
yaml_path, data = ele.data_yaml()
InIDiver = ele.InIDiver
env = ele.env
WaitClickEle = ele.WaitClickEle
ClickEle = ele.ClickEle
SendKey_Ele = ele.SendKey_Ele
GetText = ele.GetText
eleScreenshot = ele.screenshot
SwitchToFrame = ele.SwitchToFrame
SwitchToParent = ele.SwitchToParent
SwitchToDefault = ele.SwitchToDefault
GET_CurUrl = ele.GET_CurUrl
WaitSendEle = ele.WaitSendEle
SwitchWindow = ele.SwitchWindow
SwitchWindow_Url = ele.SwitchWindow_Url
CloseWindow = ele.CloseWindow
GoToWindow = ele.GoToWindow
WaitGetText = ele.WaitGetText
BackWin = ele.BackWin
SelectEle = ele.SelectEle
Allure = ele.Allure
GetElement = ele.GetElement
MaximizeWindow = ele.MaximizeWindow
GET_Title = ele.GET_Title
Enter_Url = ele.Enter_Url
GetEleExceptionEle = ele.GetEleExceptionEle
JavaScript = ele.JavaScript
Refresh = ele.Refresh
GetTimeStr = ele.GetTimeStr
JS_Click = ele.JS_Click
JS_Value = ele.JS_Value
Clear_Ele = ele.Clear_Ele
actions_click = ele.actions_click
actions_move = ele.actions_move
actions_move_click = ele.actions_move_click
actions_by_offset_click = ele.actions_by_offset_click
Screenshot = ele.Screenshot
is_displayed = ele.is_displayed
is_enabled = ele.is_enabled
is_selected = ele.is_selected
New_openpage = ele.New_openpage
switch_to_alert = ele.switch_to_alert