from selenium.webdriver.common.by import By

from base.basepage import Basepage

class Home(Basepage):
    switch_page = (By.XPATH, '//*[@id="app"]/div[2]/main/div[2]/div/div[2]/button/span')
    bv = (By.XPATH, '//*[@id="app"]/div[2]/main/div[2]/div/div[1]/div[2]/div/div/div/a')
    def home_action(self):
        # attribute_old = self.get_attribute(self.bv,'href')
        # self.open_url('https://www.bilibili.com/')
        self.click(self.switch_page)
        # assert attribute_old != self.get_attribute(self.bv, 'href')


