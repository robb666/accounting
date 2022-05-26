
from __future__ import print_function
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from accounting.faktury_GmailAPI import zsanpl
from time import sleep
from accounting.L_H_ks import url, san_l, san_h
from random_forests_selfhealing import scrp, healed_locator
from site_elements import Elements

# print(Elements.accept_1)


def san(url):
    options = webdriver.ChromeOptions()
    preferences = {'download.default_directory': r'C:\Users\PipBoy3000\Desktop\\'}
    options.add_experimental_option("prefs", preferences)
    options.add_experimental_option('excludeSwitches', ['enable-logging'])  # win devtools supress
    driver = webdriver.Chrome(executable_path=r'M:\zzzProjekty/drivery przegądarek/chromedriver.exe',
                              options=options)
    driver.get(url)
    click = ''
    healed_locator(driver, element=Elements.accept_1, helper_attr='', value=click)
    healed_locator(driver, element=Elements.login, helper_attr='', value=click)
    sleep(.7)

    helper_attr = "and contains(@class, 'button')"
    healed_locator(driver, element=Elements.san, helper_attr=helper_attr, value='')
    WebDriverWait(driver, 1).until(EC.url_changes(url))

    helper_attr = "or contains(text(), 'Akceptuję')"
    healed_locator(driver, element=Elements.accept_2, helper_attr=helper_attr, value='')

    healed_locator(driver, element=Elements.nik, helper_attr='',  value=san_l)
    healed_locator(driver, element=Elements.button_nik, helper_attr='',  value='')
    sleep(1)
    healed_locator(driver, element=Elements.ordinarypin, helper_attr='',  value=san_h)
    healed_locator(driver, element=Elements.button_ordinarypin, helper_attr='',  value='')
    healed_locator(driver, element=Elements.oneTimeAccess, helper_attr='',  value='')
    # sleep(1800)

    sleep(3.5)

    tiktok = zsanpl()
    healed_locator(driver, element=Elements.otp, helper_attr='',  value=tiktok)
    healed_locator(driver, element=Elements.otp_button, helper_attr='',  value='')
    try:
        sleep(1000)
    except:
        driver.quit()


if __name__ == '__main__':
    url = url
    san(url)
