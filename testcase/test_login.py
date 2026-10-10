import time
import allure
from selenium.webdriver.common.by import By


class TestLogin:
    but = (By.XPATH, '//*[@id="chat-submit-button"]')
    @allure.title('这是一个测试用例')
    def test_login(self, l_login):
        l_login.login()
        time.sleep(2)
        assert l_login.get_url() == 'https://www.bilibili.com/'
