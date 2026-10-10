from importlib import import_module

from page.login import Login
from web import Web
from base.basepage import Basepage
from page import login
from page.home import Home
import pytest
from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.chrome.service import Service as ChromeService

# @pytest.fixture(scope="session")
# def driver():
#     service = ChromeService(executable_path='./chromedriver.exe')
#     driver = webdriver.Chrome(service=service)
#     # driver.get('https://www.baidu.com/index.htm')
#     driver.maximize_window()
#     driver.implicitly_wait(10)
#     page = Basepage(driver)
#     return page
#     # driver.quit()

@pytest.fixture(scope="session")
def web_driver():
    driver = Web.start()
    yield driver
    Web.quit()

@pytest.fixture()
def l_login(web_driver):
    return Login(web_driver)

@pytest.fixture()
def h_home(web_driver):
    return Home(web_driver)

# @pytest.fixture(scope='module')
# def web_quit():
#     yield
#     Web.quit()

# @pytest.hookimpl(hookwrapper=True)
# def pytest_runtest_makereport(item, call):
#     outcome = yield
#     rep = outcome.get_result()
#     # rep.when: setup / call / teardown
#     if rep.when == "call":
#         nodeid = item.nodeid  # testcase/test_login.py::TestLogin::test_login
#         status = rep.outcome # passed / failed / skipped 小写
#         print(f"{nodeid} {status.upper()}")
#         print("详细信息：{}".format(item))
