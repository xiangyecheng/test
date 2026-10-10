import time
import allure
from selenium.webdriver.common.by import By


class TestLogin:
    but = (By.XPATH, '//*[@id="chat-submit-button"]')
    login_assert = (By.XPATH, '/html/body/div[3]/div/div[2]/div[1]')
    @allure.title('这是一个测试用例')
    def test_login(self, l_login):
        l_login.login()
        assert l_login.find_element(self.login_assert).text == '扫描二维码登录'
        time.sleep(2)