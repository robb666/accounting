
from __future__ import print_function
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from accounting.faktury_GmailAPI import zsanpl
import time
from L_H_ks import url, san_l, san_h
from random_forests_selfhealing import scrp, healed_locator


def san(url):
    options = webdriver.ChromeOptions()
    preferences = {'download.default_directory': r'C:\Users\PipBoy3000\Desktop\\'}
    options.add_experimental_option("prefs", preferences)
    options.add_experimental_option('excludeSwitches', ['enable-logging'])  # win devtools supress
    driver = webdriver.Chrome(executable_path=r'M:\zzzProjekty/drivery przegądarek/chromedriver.exe',
                              options=options)
    driver.get(url)
    healed_locator(driver, header=8, element_row=0, helper_attr='', value='')
    time.sleep(1800)
    # WebDriverWait(driver, 1).until(EC.element_to_be_clickable((
    #                                     By.ID, 'privacy-prompt-controls-button-accept'))).click()

    healed_locator(driver, header=0, element_row=0, helper_attr='', value='')

    helper_attr = "and parent::node()[@class='quick_links__text']"
    healed_locator(driver, header=2, element_row=0, helper_attr=helper_attr, value='')

    helper_attr = "or text()='Biura regionalne'"
    healed_locator(driver, header=6, element_row=0, helper_attr=helper_attr, value='')



if __name__ == '__main__':
    url = url
    san(url)
