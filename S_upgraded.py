
from __future__ import print_function
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from faktury_GmailAPI import zsanpl
from time import sleep
from accounting.L_H_ks import url, san_l, san_h
from random_forests_selfhealing import scrp, healed_locator


def san(url):
    options = webdriver.ChromeOptions()
    preferences = {'download.default_directory': r'C:\Users\PipBoy3000\Desktop\\'}
    options.add_experimental_option("prefs", preferences)
    options.add_experimental_option('excludeSwitches', ['enable-logging'])  # win devtools supress
    driver = webdriver.Chrome(executable_path=r'M:\zzzProjekty/drivery przegądarek/chromedriver.exe',
                              options=options)
    driver.get(url)
    click = ''
    healed_locator(driver, header=14, element_row=0, helper_attr='', value=click)
    sleep(1800)
    healed_locator(driver, header=0, element_row=0, helper_attr='', value=click)
    sleep(.7)

    helper_attr = "and contains(@class, 'button')"
    healed_locator(driver, header=0, element_row=1, helper_attr=helper_attr, value='')
    WebDriverWait(driver, 1).until(EC.url_changes(url))

    # helper_attr = "or contains(text(), 'Akceptuję')"
    healed_locator(driver, header=16, element_row=0, helper_attr='', value='')

    healed_locator(driver, header=3, element_row=0, helper_attr='',  value=san_l)
    healed_locator(driver, header=3, element_row=1, helper_attr='',  value='')
    sleep(1)
    healed_locator(driver, header=8, element_row=0, helper_attr='',  value=san_h)
    healed_locator(driver, header=8, element_row=1, helper_attr='',  value='')
    healed_locator(driver, header=6, element_row=0, helper_attr='',  value='')

    sleep(3.5)

    tiktok = zsanpl()
    healed_locator(driver, header=11, element_row=0, helper_attr='',  value=tiktok)
    healed_locator(driver, header=11, element_row=1, helper_attr='',  value='')
    try:
        sleep(1000)
    except:
        driver.quit()


if __name__ == '__main__':
    url = url
    san(url)
