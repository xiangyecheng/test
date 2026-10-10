import time
import allure
from base.basepage import Basepage
from selenium.webdriver.common.by import By

class Login(Basepage):
    but = (By.XPATH, '//*[@id="chat-submit-button"]')
    @allure.title('这是一个测试用例')
    def login(self):
        self.click(self.but)
        self.open_new_url('https://www.bilibili.com/')
        self.switch_windows()
        # driver.execute_script("window.open('{}');".format('https://www.bilibili.com/'))
        # Driver.get("https://www.bilibili.com/")
        # print(2/1)
