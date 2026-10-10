import time
import os
from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.chrome.service import Service as ChromeService
from logs.log import Log
logger = Log().get_log()

class Basepage:
    # service = ChromeService(executable_path='./chromedriver.exe')
    # driver = webdriver.Chrome(service=service)
    def __init__(self, driver: WebDriver):
        self.driver = driver
        self.g = globals()
        self.current_test = os.environ.get('PYTEST_CURRENT_TEST')

    def open_url(self, url):
        # self.service = ChromeService(executable_path='./chromedriver.exe')
        # self.driver = webdriver.Chrome(service=self.service)
        self.driver.get(url)
        self.driver.maximize_window()
        logger.info('启动浏览器，最大化')

    def open_new_url(self,url):
        self.driver.execute_script("window.open('{}');".format(url))
        logger.info('{}打开url：{}'.format(self.current_test, url))

    def get_url(self):
        """获取url"""
        try:
            url = self.driver.current_url
            logger.info('{}获取到地址：{}'.format(self.current_test, url))
            return url
        except Exception as err:
            logger.error('{}获取地址失败{}'.format(self.current_test, err))

    def quit(self):
        self.driver.quit()
        logger.info('退出浏览器')

    def click(self, element):
        try:
            self.driver.find_element(*element).click()
            logger.info('{}点击元素：{}'.format(self.current_test,element))
        except Exception as err:
            logger.error('{}点击元素失败，元素：{}'.format(self.current_test,element))

    def switch_windows(self):
        handles = self.driver.window_handles
        self.driver.switch_to.window(handles[-1])
