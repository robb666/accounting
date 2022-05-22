
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

    WebDriverWait(driver, 1).until(EC.element_to_be_clickable((
                                        By.ID, 'privacy-prompt-controls-button-accept'))).click()
    healed_locator(driver, e=None, attr=None, header=0, helper_attr='', element_row=0, value='')
    time.sleep(.7)

    helper_attr = "and contains(@class, 'button')"
    healed_locator(driver, e=None, attr=None, header=0,  helper_attr=helper_attr, element_row=1, value='')
    WebDriverWait(driver, 1).until(EC.url_changes(url))
    try:
        driver.find_element_by_xpath('//div[contains(@id, "button-accept")]').click()
    except:
        pass
    healed_locator(driver, e=None, attr=None, helper_attr='', header=3, element_row=0, value=san_l)
    healed_locator(driver, e=None, attr=None, helper_attr='', header=3, element_row=1, value='')
    time.sleep(1)
    healed_locator(driver, e=None, attr=None, helper_attr='', header=8, element_row=0, value=san_h)
    healed_locator(driver, e=None, attr=None, helper_attr='', header=8, element_row=1, value='')
    healed_locator(driver, e=None, attr=None, helper_attr='', header=6, element_row=0, value='')

    time.sleep(3.5)

    tiktok = zsanpl()
    healed_locator(driver, e=None, attr=None, helper_attr='', header=11, element_row=0, value=tiktok)
    healed_locator(driver, e=None, attr=None, helper_attr='', header=11, element_row=1, value='')
    try:
        time.sleep(1000)
    except:
        driver.quit()


if __name__ == '__main__':
    url = url
    san(url)
