import time
import allure
from base.basepage import Basepage
from selenium.webdriver.common.by import By

class Login(Basepage):
    login_but = (By.XPATH, '//*[@id="app"]/div[2]/div[1]/div[1]/div[3]/div/div[1]/div/div[1]/div')
    @allure.title('这是一个测试用例')
    def login(self):
        self.click(self.login_but)
        # driver.execute_script("window.open('{}');".format('https://www.bilibili.com/'))
        # Driver.get("https://www.bilibili.com/")
        # print(2/1)
