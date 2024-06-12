import pytest

from selenium import webdriver
# ChromeService
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options as ChromeOptions
# EdgeService
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.options import Options as EdgeOptions
# FirefoxService
from selenium.webdriver.firefox.service import Service as FirefoxService
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.firefox.options import Options as FirefoxOptions

from .Module import *

class Base:
    @pytest.fixture
    def newDriver(self):
        op = ChromeOptions()
        # op = EdgeOptions()
        # op = FirefoxOptions()
        current_dir = os.getcwd() # 獲取當前工作目錄
        prefs = {"download.default_directory" : os.path.abspath(current_dir + r'/tests/Web/files')} # 指定到某一個下載路徑
        op.add_experimental_option("prefs", prefs) # 加入額外設定參數
        # op.set_preference('browser.download.dir', os.path.abspath(r'.\tests\Web\files')) # 指定到某一個下載路徑 Firefox
        # op.add_argument("--incognito") # 設定成以無痕視窗開啟瀏覽器 Chrome
        # op.add_argument("--inprivate") # 設定成以無痕視窗開啟瀏覽器 Edge
        # op.add_argument("--private") # 設定成以無痕視窗開啟瀏覽器 Firefox
        # op.add_argument("--headless") # 不開啟實體瀏覽器在背景執行
        driver = webdriver.Remote(command_executor='http://localhost:4449/wd/hub', options=op)
        # driver = webdriver.Chrome(service=ChromeService(executable_path=os.path.abspath(r'.\tests\Web\chromedriver.exe')), options=op)
        # driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=op)
        # driver = webdriver.Firefox(service=FirefoxService(GeckoDriverManager().install()), options=op)
        driver.maximize_window()
        driver.switch_to.window(driver.current_window_handle) # 切換到當前視窗
        InIDiver(driver)
        GoToWindow(env('mms_ordermanagement'))
        yield driver # Case 結束後執行以下命令
        driver.quit() # 關閉 driver
        return driver # 因為有多個 Case 要跑，所以要 return driver，否則跑下一個 Case 就會找不到 driver
