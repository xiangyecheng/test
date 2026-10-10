import time
import allure
from selenium.webdriver.common.by import By


class TestLogin:
    but = (By.XPATH, '//*[@id="chat-submit-button"]')
    @allure.title('这是一个测试用例')
    def test_login(self, web_driver):
        web_driver.login()
        time.sleep(2)
        assert web_driver.get_url() == 'https://www.bilibili.com/'
