
from __future__ import print_function
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.firefox.options import Options
from accounting.faktury_GmailAPI import zsanpl
import time
import os
import pickle
import os.path
from googleapiclient.discovery import build
import base64
import mimetypes
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders
from datetime import datetime
from dateutil.relativedelta import relativedelta
from L_H_ks import url, san_l, san_h
from random_forests_selfhealing import healed_locator


def san(url):
    options = webdriver.ChromeOptions()
    preferences = {'download.default_directory': r'C:\Users\PipBoy3000\Desktop\\'}
    options.add_experimental_option("prefs", preferences)
    options.add_experimental_option('excludeSwitches', ['enable-logging'])  # win devtools supress
    driver = webdriver.Chrome(executable_path=r'M:\zzzProjekty/drivery przegądarek/chromedriver.exe',
                              options=options)
    driver.get(url)
    time.sleep(.7)
    try:
        driver.find_element_by_id('privacy-prompt-controls-button-accept').click()
    except:
        pass
    driver.find_element_by_xpath('//span[contains(text(), "Zaloguj")]').click()
    time.sleep(.7)



    try:
        WebDriverWait(driver, 1).until(EC.presence_of_element_located((
            # By.XPATH, "//a[contains(@href, 'centrum24-web/login') and contains(@class, 'button')]"))).click()
            By.XPATH, "//*[contains(text(), 'Santander internet')]"))).click()
    except Exception as e:
        print('Exc messa.', e)
        helper_attr = "and contains(@class, 'button')"
        healed_locator(driver, e,
                       attr='title', header=0,  helper_attr=helper_attr, element_row=0, value='', filename='Test.csv')


    WebDriverWait(driver, 1).until(EC.url_changes(url))


    try:
        driver.find_element_by_xpath('//div[contains(@id, "button-accept")]').click()
    except:
        pass


    try:
        login = driver.find_element_by_id('input_nikko')
        WebDriverWait(driver, 1).until(EC.presence_of_element_located((
            # By.XPATH, "//*[contains(@href, 'centrum24-web/login')]"))).click()
            By.XPATH, "//a[contains(@href, 'centrum24-web/login') and contains(@class, 'button')]"))).click()
        login.send_keys(san_l)
        time.sleep(1.3)
    except Exception as e:
        print('Exc messa. login', e)
        # helper_attr = "contains(@class, 'button')"
        healed_locator(driver, e,
                       attr='id', helper_attr='', header=2, element_row=0, value='100200', filename='Test.csv')





    try:
        WebDriverWait(driver, 9).until(EC.presence_of_element_located((By.XPATH, "//input[@id='okBtn2']"))).click()
    except:
        WebDriverWait(driver, 9).until(EC.presence_of_element_located((By.XPATH, "//input[@id='okBtn2']"))).click()

    time.sleep(1.3)

    pwd = driver.find_element_by_id('ordinarypin')
    pwd.send_keys(san_h)
    driver.find_element_by_id('okBtn2').click()

    onet = driver.find_element_by_id('back-button')
    onet.click()
    time.sleep(3.5)

    tiktok = zsanpl()
    driver.find_element_by_id('input_nik').send_keys(tiktok)
    driver.find_element_by_id('okBtn2').click()
    try:
        time.sleep(1000)
    except:
        driver.quit()


if __name__ == '__main__':
    url = url
    san(url)
