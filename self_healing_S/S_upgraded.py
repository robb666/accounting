
from __future__ import print_function
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from faktury_GmailAPI import zsanpl
from time import sleep
from L_H_ks import url, san_l, san_h
from random_forests_selfhealing import healed_locator
from site_elements import Elements


def san(url):
    options = webdriver.ChromeOptions()
    preferences = {'download.default_directory': r'C:\Users\PipBoy3000\Desktop\\'}
    options.add_experimental_option("prefs", preferences)
    options.add_experimental_option('excludeSwitches', ['enable-logging'])  # win devtools supress
    driver = webdriver.Chrome(executable_path=r'M:\zzzProjekty/drivery przegądarek/chromedriver.exe',
                              options=options)

    driver.get(url); click = ''

    helper_attr = "or contains(text(), 'Akceptuję')"
    healed_locator(driver, element=Elements.accept_2, helper_attr=helper_attr, value=click)
    healed_locator(driver, element=Elements.nik, value=san_l)
    healed_locator(driver, element=Elements.button_nik, value=click)
    sleep(1)

    healed_locator(driver, element=Elements.ordinarypin, value=san_h)
    healed_locator(driver, element=Elements.button_ordinarypin, value=click)
    healed_locator(driver, element=Elements.oneTimeAccess, value=click)
    sleep(3.5)

    tiktok = zsanpl()
    healed_locator(driver, element=Elements.otp, value=tiktok)
    healed_locator(driver, element=Elements.otp_button, value=click)
    try:
        sleep(1000)
    except:
        driver.quit()


if __name__ == '__main__':
    url = url
    san(url)
