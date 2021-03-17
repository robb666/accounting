import os
import multiprocessing
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.firefox.options import Options
from faktury_GmailAPI import main
from L_H_ks import san_l, san_h, allianz_l, allianz_h, compensa_l, compensa_h, generali_l, generali_h, \
     hestia_l, hestia_h, uniqa_l, uniqa_h, warta_l, warta_h, interrisk_l, interrisk_h, proama_l, proama_h, \
     unilink_l, unilink_h, pzu_l, pzu_h, warta_ż_l, warta_ż_h
import time





### SANTANDER wyciąg ###
def santander():
    ### CHROME ###
    options = webdriver.ChromeOptions()
    preferences = {'download.default_directory': "C:\\Users\\ROBERT\\Desktop\\Księgowość\\2019"}
    options.add_experimental_option("prefs", preferences)
    driver = webdriver.Chrome(executable_path=r'M:/zzzProjekty/drivery przegądarek/chromedriver.exe', options=options)

    try:
        url_santander = 'https://santander.pl/'
        driver.get(url_santander)
        driver.find_element_by_partial_link_text('Zaloguj').click()
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Santander interne"))).click()
        try:
            driver.switch_to.window(driver.window_handles[1])
        except:
            pass
        login_san = driver.find_element_by_id('input_nik')
        login_san.send_keys(san_l)
        time.sleep(2)
        dalej = driver.find_element_by_id('okBtn2')
        dalej.click()
        hasło_san = WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.ID, "ordinarypin")))
        hasło_san.send_keys(san_h)
        driver.find_element_by_id('okBtn2').click()
        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'favourite-element'))).click()
        driver.find_element_by_partial_link_text("Pobie").click()
        time.sleep(2)
        driver.find_element_by_class_name('logout').click()
        driver.quit()
        print('Santander ok')

    except:
        print('Brak wyciągu bankowego')
        driver.quit()
        pass



def allianz():
    options = webdriver.ChromeOptions()
    preferences = {'download.default_directory': "C:\\Users\\ROBERT\\Desktop\\Księgowość\\2021\\RobO"}
    options.add_experimental_option("prefs", preferences)
    driver = webdriver.Chrome(executable_path=r'M:/zzzProjekty/drivery przegądarek/chromedriver.exe', options=options)

    try:
        url_allianz = 'https://start.allianz.pl/'
        driver.get(url_allianz)
        login = driver.find_element_by_id('username')
        login.send_keys(allianz_l)
        hasło = driver.find_element_by_id('password')
        hasło.send_keys(allianz_h)
        driver.find_element_by_name('submit').click()
        driver.get('https://chuck.allianz.pl/agent/#/invoices')
        WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.am-btn-large'))).click()
        time.sleep(0.8)
        WebDriverWait(driver, 7).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.am-btn-primary')))[1].click()
        time.sleep(4)
        driver.quit()
        print('Allianz ok')

    except:
        time.sleep(1)
        print()
        print('Brak Allianz')
        driver.quit()
        pass


def compensa():
    try:
        driver_F = webdriver.Chrome(executable_path=r'M:/zzzProjekty/drivery przegądarek/chromedriver.exe')#, options=options)
        url_compensa = 'https://cportal.compensa.pl/'
        driver_F.get(url_compensa)
        WebDriverWait(driver_F, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.input'))).send_keys(compensa_l)
        driver_F.find_element_by_css_selector('div.fl:nth-child(5) > input:nth-child(1)').send_keys(compensa_h)
        WebDriverWait(driver_F, 5).until(EC.presence_of_element_located((By.ID, "btnLogin"))).click()
        WebDriverWait(driver_F, 5).until(EC.presence_of_element_located((By.ID, 'News')))
        url_compensa = 'https://cportal.compensa.pl/#MyCommissions'
        driver_F.get(url_compensa)
        driver_F.set_page_load_timeout(40)
        driver_F.get_screenshot_as_file('C:/Users/ROBERT/Desktop/Księgowość/2021/RobO/compensa.png')
        print('Compensa ok')

    except Exception as e:
        print('Brak Compensa')


def generali():
    options = webdriver.ChromeOptions()
    preferences = {'download.default_directory': "C:\\Users\\ROBERT\\Desktop\\Księgowość\\2021\\RobO"}
    options.add_experimental_option("prefs", preferences)
    driver = webdriver.Chrome(executable_path=r'M:/zzzProjekty/drivery przegądarek/chromedriver.exe', options=options)
    try:
        url_generali = 'https://portal.generali.pl/auth/login?service=https%3A%2F%2Fportal.generali.pl%2Flogin%2Fcas'
        driver.get(url_generali)
        WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#username'))).send_keys(generali_l)
        driver.find_element_by_css_selector('#password').send_keys(generali_h)
        driver.find_element_by_css_selector('#fm1 > div.login > input[type=submit]:nth-child(6)').click()
        url_accounting = 'https://portal.generali.pl/mikado/commissions/current'
        driver.get(url_accounting)
        WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.XPATH, "//*[@class='far fa-file-zip-o']"))).click()
        time.sleep(9)
        driver.close()
        print('Generali ok')
    except:
        print('Brak Generali')
        driver.close()
        pass


def hestia():
    options = webdriver.ChromeOptions()
    preferences = {'download.default_directory': "C:\\Users\\ROBERT\\Desktop\\Księgowość\\2021\\RobO"}
    options.add_experimental_option("prefs", preferences)
    driver = webdriver.Chrome(executable_path=r'M:/zzzProjekty/drivery przegądarek/chromedriver.exe', options=options)
    try:
        url_santander = 'https://sso.ergohestia.pl/my.policy'
        driver.get(url_santander)
        login_hes = driver.find_element_by_id('input_1').send_keys(hestia_l)
        hasło_hes = driver.find_element_by_id('input_2').send_keys(hestia_h)
        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH,
                                                                '//*[@id="auth_form"]/div[3]/div[4]/button'))).click()
        url_agent = 'https://partner.ergohestia.pl/#/partner'
        driver.get(url_agent)
        url_agent = 'https://partner.ergohestia.pl/#/partner' + '/commissionHistory'
        driver.get(url_agent)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,
                                                                        '/html/body/div/div[3]/div[2]/div/ng-include/'
                                                                        'div/div[3]/div/div/div/div[2]/table/tbody/'
                                                                        'tr[1]/td[6]/ng-include/a'))).click()
        time.sleep(1.2)
        faktura = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,
                                                                        '/html/body/div/div[3]/div[2]/div/ul/li[2]/a')))
        faktura.click()
        time.sleep(2)
        driver.quit()
        print('Hestia ok')
    except:
        print('Brak Hestii')
        driver.quit()




def interrisk() :
    options = webdriver.ChromeOptions()
    preferences = {'download.default_directory' : "C:\\Users\\ROBERT\\Desktop\\Księgowość\\2021\\RobO"}
    options.add_experimental_option("prefs", preferences)
    driver = webdriver.Chrome(executable_path=r'M:/zzzProjekty/drivery przegądarek/chromedriver.exe', options=options)
    try :
        url_interrisk = 'https://portal.interrisk.pl/Zaloguj'
        driver.get(url_interrisk)
        driver.find_element_by_id('ctl00_cph1_uxLogin_UserName').send_keys(interrisk_l)
        WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.ID, "ctl00_cph1_uxLogin_Password"))).send_keys(interrisk_h)
        WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.ID, "ctl00_cph1_uxLogin_LoginButton"))).click()

        url_interrisk_prow = 'https://portal.interrisk.pl/Rozliczenia/NotyProwizyjne/Przegladaj'
        driver.get(url_interrisk_prow)
        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#ctl00_ctl00_cph1_cph1_search"))).click()
        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[3]/div[2]/div[1]/div[2]/div[2]/table/tbody/tr[1]/td[6]/input'))).click()
        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, 'ctl00_ctl00_cph1_cph1_cbNoteOnDemand'))).click()
        time.sleep(1)
        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#ctl00_ctl00_cph1_cph1_exportPdf"))).click()
        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#ctl00_ctl00_cph1_cph1_search"))).click()
        try :
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[3]/div[2]/div[1]/div[2]/div[2]/table/tbody/tr[2]/td[6]/input'))).click()
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#ctl00_ctl00_cph1_cph1_exportPdf"))).click()
        except :
            pass
        time.sleep(1)
        driver.quit()
        print('InterRisk ok')
    except :
        driver.quit()
        print('Brak InterRisk')



### UNIQA ###
def uniqa():
    options = webdriver.ChromeOptions()
    preferences = {'download.default_directory': "C:\\Users\\ROBERT\\Desktop\\Księgowość\\2021\\RobO"}
    options.add_experimental_option("prefs", preferences)
    driver = webdriver.Chrome(executable_path=r'M:/zzzProjekty/drivery przegądarek/chromedriver.exe', options=options)
    payload = {'login': uniqa_l,
               'password': uniqa_h}
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.86 Safari/537.36'}
    try:
        with requests.Session() as s:
            url_post_login = 'https://pos.uniqa.pl/pl/login_fe'
            r = s.post(url_post_login, data=payload, headers=headers)
            ks = s.get('https://pos.uniqa.pl/pl/zadania_i_plany/prowizje?menu=1')
        driver.get(ks.url)
        driver.delete_all_cookies()

        for cookie in s.cookies.items() :
            driver.add_cookie({"name": cookie[0], "value" : cookie[1]})
        driver.get(ks.url)

        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#form2\:showFilem2"))).click()
        time.sleep(3)
        driver.quit()
        print('Uniqa ok')
    except:
        driver.quit()
        print('Brak UNIQA')
        pass




def warta():
    options = webdriver.ChromeOptions()
    preferences = {'download.default_directory' : "C:\\Users\\ROBERT\\Desktop\\Księgowość\\2021\\RobO"}
    options.add_experimental_option("prefs", preferences)
    driver = webdriver.Chrome(executable_path=r'M:/zzzProjekty/drivery przegądarek/chromedriver.exe', options=options)
    try :
        url_warta = 'https://cas.warta.pl/cas/login?service=https%3A%2F%2Feagent.warta.pl%2Fview360%2Flogin%2Fcas'
        driver.get(url_warta)
        driver.find_element_by_id('username').send_keys(warta_l)
        driver.find_element_by_id('password').send_keys(warta_h)
        driver.find_element_by_name('submit').click()
        try :
            if driver.find_element_by_name('continue') != 0 :
                driver.find_element_by_name('continue').click()
        except :
            pass
        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH,
                                                                       "//*[contains(text(), 'Majątek')]"))).click()
        time.sleep(0.9)
        rozliczenia_agencji = 'https://eagent.warta.pl/view360/#/app/main/settlement/property/A00005152001/agent/list/?aid=1770143&agentOuid=A00005152001'
        driver.get(rozliczenia_agencji)
        time.sleep(1.1)
        WebDriverWait(driver, 4).until(EC.presence_of_all_elements_located((By.XPATH, "//*[contains(text(), 'RSP')]")))[0].click()
        WebDriverWait(driver, 4).until(EC.presence_of_all_elements_located((By.CLASS_NAME, "settlement-details-documents__content__item__list__elem")))[2].click()
        driver.get(rozliczenia_agencji)
        time.sleep(0.9)
        WebDriverWait(driver, 4).until(EC.presence_of_all_elements_located((By.XPATH, "//*[contains(text(), 'RSP')]")))[1].click()
        WebDriverWait(driver, 4).until(EC.presence_of_all_elements_located((By.CLASS_NAME, "settlement-details-documents__content__item__list__elem")))[2].click()
        time.sleep(2)
        driver.quit()
        print('Warta ok')
    except:
        driver.quit()
        print('Brak WARTA')
        pass


def warta_ż():
    options = webdriver.ChromeOptions()
    preferences = {'download.default_directory': "C:\\Users\\ROBERT\\Desktop\\Księgowość\\2021\\RobO",
                   'plugins.always_open_pdf_externally': True}
    options.add_experimental_option("prefs", preferences)
    driver = webdriver.Chrome(executable_path=r'M:/zzzProjekty/drivery przegądarek/chromedriver.exe', options=options)
    try:
        url_warta_ż = 'https://eplatforma.warta.pl/'
        driver.get(url_warta_ż)
        WebDriverWait(driver, 4).until(EC.presence_of_element_located((By.NAME, "LOGNAME_13"))).send_keys(warta_ż_l)
        WebDriverWait(driver, 4).until(EC.presence_of_element_located((By.NAME, "PASSWD_13"))).send_keys(warta_ż_h)
        WebDriverWait(driver, 4).until(EC.presence_of_element_located((By.NAME, "zaloguj"))).click()
        WebDriverWait(driver, 4).until(
            EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'samorozliczenie')]"))).click()
        WebDriverWait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR,
                                                                       ".filtr > th:nth-child(8) > input:nth-child(1)"))).click()
        WebDriverWait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR,
                                                                       'tr.td_line2:nth-child(4) > td:nth-child(8) > a:nth-child(1)'))).click()
        time.sleep(9)
        driver.quit()
        print('Warta Ż ok')
    except:
        driver.quit()
        print('Brak Warta Ż')




### UNILINK ###
def unilink():
    ### CHROME ###
    options = webdriver.ChromeOptions()
    preferences = {'download.default_directory': "C:\\Users\\ROBERT\\Desktop\\Księgowość\\2021\\RobO"}
    options.add_experimental_option("prefs", preferences)
    driver = webdriver.Chrome(executable_path=r'M:/zzzProjekty/drivery przegądarek/chromedriver.exe', options=options)

    try:
        url_unilink = 'https://unilink.pl/logowanie'
        driver.get(url_unilink)

        WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.ID, "login"))).send_keys(unilink_l)
        WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.ID, "password"))).send_keys(unilink_h)
        WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.ID, "submit"))).click()
        time.sleep(1.4)
        url_unilink_faktury = 'https://unilink.pl/pokaz/4020'
        driver.get(url_unilink_faktury)
        WebDriverWait(driver, 7).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="searchformlist"]/table/tbody/tr[2]/td[13]/div[1]/a/i'))).click()
        time.sleep(3)
        driver.quit()
        print('Unilink ok')
    except:
        driver.quit()
        print('Brak Unilink')
        pass




### PZU ###
def pzu():
    ### CHROME ###
    options = webdriver.ChromeOptions()
    preferences = {'download.default_directory': "C:\\Users\\ROBERT\\Desktop\\Księgowość\\2021\\RobO"}
    options.add_experimental_option("prefs", preferences)
    driver = webdriver.Chrome(executable_path=r'M:/zzzProjekty/drivery przegądarek/chromedriver.exe', options=options)

    try:
        driver.get('https://everest.pzu.pl/pc/PolicyCenter.do')
        login = driver.find_element_by_id('input_1')
        login.send_keys(pzu_l)
        hasło = driver.find_element_by_id('input_2')
        hasło.send_keys(pzu_h)
        driver.find_element_by_css_selector('.credentials_input_submit').click()
        login = driver.find_element_by_id('Login:LoginScreen:LoginDV:username-inputEl')
        login.send_keys(pzu_l)
        hasło = driver.find_element_by_id('Login:LoginScreen:LoginDV:password-inputEl')
        hasło.send_keys(pzu_h)
        driver.find_element_by_id('Login:LoginScreen:LoginDV:submit').click()
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, 'treeview-1059-record-9'))).click()
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, 'ProducerStatementReportOnlinePzu:0:statementTab-btnInnerEl'))).click()
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, 'ProducerStatementReportOnlinePzu:0:StatementsLV:1:DownloadPdfFileLink'))).click()
        time.sleep(2)
        driver.quit()
        print('PZU ok')

    except:
        driver.quit()
        print('Brak PZU')
        pass


# santander()
# allianz()
# compensa()
# generali()
# hestia()
# interrisk()
# uniqa()
# warta()
# warta_ż()
unilink()
# pzu()









# if __name__ == '__main__':
#     # multiprocessing.freeze_support()
#
#     p1 = multiprocessing.Process(target=santander)
#     p2 = multiprocessing.Process(target=allianz)
#     p3 = multiprocessing.Process(target=main)        ### AXA | WIENER | INSLY ###
#     p4 = multiprocessing.Process(target=compensa)
#     p5 = multiprocessing.Process(target=generali)
#     p6 = multiprocessing.Process(target=hestia)
#     p7 = multiprocessing.Process(target=interrisk)
#     p8 = multiprocessing.Process(target=proama)
#     p9 = multiprocessing.Process(target=uniqa)
#     p10 = multiprocessing.Process(target=warta)
#     p11 = multiprocessing.Process(target=warta_ż)
#     p12 = multiprocessing.Process(target=unilink)
#     p13 = multiprocessing.Process(target=pzu)
#
# ##################
#
#     p1.start()
#     p2.start()
#     p3.start()
#
#     p1.join()
#     p2.join()
#     p3.join()
#
# ###################
#
#     p4.start()
#     p5.start()
#     p6.start()
#
#     p4.join()
#     p5.join()
#     p6.join()
#
# ####################
#
#     p7.start()
#     p8.start()
#     p9.start()
#
#     p7.join()
#     p8.join()
#     p9.join()
#
# ####################
#
#     p10.start()
#     p11.start()
#
#     p10.join()
#     p11.join()
#
# ####################
#
#     p12.start()
#     p13.start()
#
#     p12.join()
#     p13.join()
