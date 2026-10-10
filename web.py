import time
# from base.basepage import Basepage
from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.chrome.service import Service as ChromeService
from page.login import Login
import os
import allure

class Web:
    url = 'https://www.baidu.com/index.htm'

    service = ChromeService(executable_path='./chromedriver.exe')
    driver = webdriver.Chrome(service=service)

    @classmethod
    def start(cls):
        # cls.service = ChromeService(executable_path='./chromedriver.exe')
        # cls.driver = webdriver.Chrome(service=cls.service)
        # cls.driver.get('https://www.baidu.com/index.htm')
        # cls.open_url('https://www.baidu.com/index.htm')
        # Basepage().open_url('https://www.baidu.com/index.htm')
        cls.driver.get(cls.url)
        cls.driver.maximize_window()
        cls.driver.implicitly_wait(10)
        return cls.driver

    @classmethod
    def quit(cls):
        cls.driver.quit()

    @classmethod
    def capture(cls, description=None):
        '''
        页面截屏保存截图
        :param description: 截图描述
        :return:
        '''
        current_datetime = time.strftime(r"%Y%m%d%H%M%S", time.localtime(time.time()))
        current_test = os.environ.get('PYTEST_CURRENT_TEST').split(':')[-1].split(
            ' ')[0]
        screenshot_dir = os.path.join(os.getcwd(), 'screenshot/%s' % current_test)
        if not os.path.exists(screenshot_dir):
            os.makedirs(screenshot_dir)
        filename = os.path.join(screenshot_dir, '{}.png'.format(current_datetime))
        cls.driver.get_screenshot_as_file(filename)
        with open(filename, mode='rb') as f:
            file = f.read()
        allure.attach(file, description, allure.attachment_type.PNG)

