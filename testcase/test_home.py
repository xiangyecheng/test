import time
import allure
from selenium.webdriver.common.by import By

class TestHome:
    bv = (By.XPATH, '//*[@id="app"]/div[2]/main/div[2]/div/div[1]/div[3]/div/div/div/a')
    def test_home(self, h_home):
        attribute_old = h_home.get_attribute(self.bv, 'href')
        h_home.home_action()
        time.sleep(2)
        assert attribute_old != h_home.get_attribute(self.bv, 'href')
