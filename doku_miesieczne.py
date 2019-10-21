import multiprocessing
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.firefox.options import Options
from AXA_i_Wienner_GmailAPI import main
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



### ALLIANZ ###
def allianz():
    ### CHROME ###
    options = webdriver.ChromeOptions()
    preferences = {'download.default_directory': "C:\\Users\\ROBERT\\Desktop\\Księgowość\\2019"}
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
        time.sleep(3)
        print('Allianz ok')
        # pyautogui.press(['enter'])
        # driver.quit()

    except:
        time.sleep(1)
        print()
        print('Brak Allianz')
        driver.quit()
        pass



###################
### AXA, WIENER ###
###################



### COMPENSA ###
def compensa():
    ### FIREFOX ###
    options = Options()
    options.set_preference("browser.download.folderList", 2)
    options.set_preference("browser.download.manager.showWhenStarting", False)
    options.set_preference("browser.download.dir", "C:\\Users\\ROBERT\\Desktop\\Księgowość\\2019")
    options.set_preference("browser.helperApps.neverAsk.saveToDisk",
                           "application/octet-stream,application/vnd.ms-excel")
    driver_F = webdriver.Firefox(executable_path=r'M:/zzzProjekty/drivery przegądarek/geckodriver.exe', options=options)

    try:
        url_compensa = 'https://cportal.compensa.pl/'
        driver_F.get(url_compensa)

        WebDriverWait(driver_F, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.input'))).send_keys(compensa_l)
        driver_F.find_element_by_css_selector('div.fl:nth-child(5) > input:nth-child(1)').send_keys(compensa_h)
        time.sleep(0.8)
        WebDriverWait(driver_F, 5).until(EC.presence_of_element_located((By.ID, "btnLogin"))).click()
        try:
            tuba_pay = WebDriverWait(driver_F, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'button.button:nth-child(1)')))
            if tuba_pay != 0 :
                tuba_pay.click()
        except:
            pass

        WebDriverWait(driver_F, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.button'))).click()
        url_comission = 'https://cportal.compensa.pl/#MyCommissions'
        driver_F.get(url_comission)
        try :
            remont = WebDriverWait(driver_F, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'button.button:nth-child(1)')))
            if remont != 0 :
                remont.click()
        except :
            pass

        try:
            WebDriverWait(driver_F, 3).until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'btnMore')))[0].click()
            WebDriverWait(driver_F, 4).until(EC.presence_of_element_located((By.CLASS_NAME, 'btnMore'))).click()
            print('Compensa ok')
        except:
            url_comission = 'https://cportal.compensa.pl/#MyCommissions'
            driver_F.get(url_comission)

            WebDriverWait(driver_F, 3).until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'btnMore')))[0].click()
            WebDriverWait(driver_F, 4).until(EC.presence_of_element_located((By.CLASS_NAME, 'btnMore'))).click()
            print('Compensa ok, ale po błędzie')
        WebDriverWait(driver_F, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'td.ta-r:nth-child(2)')))
        time.sleep(3)
        driver_F.save_screenshot("C:/Users/ROBERT/Desktop/Księgowość/2019/compensa.png")
        time.sleep(2)
        driver_F.close()

    except:
        print('Brak Compensa')
        driver_F.close()
        pass



### GENERALI ###
def generali():

    ### CHROME ###
    options = webdriver.ChromeOptions()
    preferences = {'download.default_directory': "C:\\Users\\ROBERT\\Desktop\\Księgowość\\2019"}
    options.add_experimental_option("prefs", preferences)
    driver = webdriver.Chrome(executable_path=r'M:/zzzProjekty/drivery przegądarek/chromedriver.exe', options=options)

    try:

        url_generali = 'https://portal.generali.pl/auth/login?service=https%3A%2F%2Fportal.generali.pl%2Flogin%2Fcas'
        driver.get(url_generali)
        WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#username'))).send_keys(generali_l)
        driver.find_element_by_css_selector('#password').send_keys(generali_h)
        driver.find_element_by_css_selector('.login > input:nth-child(4)').click()
        url_accounting = 'https://portal.generali.pl/pipp/accountingDocs.do'
        driver.get(url_accounting)
        WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.NAME, 'reportDate'))).send_keys(Keys.DOWN)



        lista = driver.find_elements_by_xpath("//table[@class='detailsGenerali']")

        if (len(lista) > 0):
            btns = driver.find_elements_by_partial_link_text("Faktura")
            for btn in btns:
                btn.click()
            time.sleep(2)
        else:
            driver.find_element_by_name('reportDate').send_keys(Keys.DOWN)
            btns = driver.find_elements_by_partial_link_text("Faktura")
            for btn in btns:
                btn.click()
            time.sleep(2)

        print('Generali ok')

    except:
        print('Brak Generali')
        driver.close()
        pass



### HESTIA ###
def hestia():
    options = webdriver.ChromeOptions()
    preferences = {'download.default_directory': "C:\\Users\\ROBERT\\Desktop\\Księgowość\\2019"}
    options.add_experimental_option("prefs", preferences)
    driver = webdriver.Chrome(executable_path=r'M:/zzzProjekty/drivery przegądarek/chromedriver.exe', options=options)
                                                # 'C:/Users/ROBERT/Desktop/IT/PYTHON/PYTHON 37 PROJEKTY/web scrapping/chromedriver.exe'
    try:
        url_santander = 'https://sso.ergohestia.pl/my.policy'
        driver.get(url_santander)
        login_hes = driver.find_element_by_id('input_1').send_keys(hestia_l)
        hasło_hes = driver.find_element_by_id('input_2').send_keys(hestia_h)
        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH,
                                                                       '//*[@id="auth_form"]/div[3]/div[4]/button'))).click()

        # time.sleep(1)
        url_agent = 'https://partner.ergohestia.pl/#/partner'
        driver.get(url_agent)
        url_agent = 'https://partner.ergohestia.pl/#/partner' + '/commissionHistory'
        driver.get(url_agent)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[3]/div[2]/div/ng-include/div/div[3]/div/div/div/div[2]/table/tbody/tr[1]/td[6]/ng-include/a'))).click()
        time.sleep(1.2)
        faktura = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[3]/div[2]/div/ul/li[2]/a')))

        faktura.click()
        time.sleep(2)
        driver.quit()
        print('Hestia ok')


    except:
        print('Brak Hestii')
        driver.quit()



### INTERRISK ###
def interrisk() :
    ### CHROME ###
    options = webdriver.ChromeOptions()
    preferences = {'download.default_directory' : "C:\\Users\\ROBERT\\Desktop\\Księgowość\\2019"}
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




### PROAMA ###
def proama():

    ### CHROME ###
    options = webdriver.ChromeOptions()
    preferences = {'download.default_directory': "C:\\Users\\ROBERT\\Desktop\\Księgowość\\2019"}
    options.add_experimental_option("prefs", preferences)
    driver = webdriver.Chrome(executable_path=r'M:/zzzProjekty/drivery przegądarek/chromedriver.exe', options=options)

    try:
        url_proama = 'https://proagent.proama.pl/'
        driver.get(url_proama)
        WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, 'ZALOGUJ'))).click()
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#username'))).send_keys(proama_l)
        driver.find_element_by_css_selector('#password').send_keys(proama_h)
        driver.find_element_by_css_selector('.login > input:nth-child(4)').click()
        url_accounting = 'https://portal.proama.pl/pipp/accountingDocs.do'
        driver.get(url_accounting)

        WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.NAME, 'reportDate'))).send_keys(Keys.DOWN)

        lista = driver.find_elements_by_xpath("//table[@class='detailsGenerali']")
        btns = driver.find_elements_by_partial_link_text("Faktura")

        if len(lista) > 0:
            for btn in btns:
                btn.click()
            time.sleep(3)
            print('Proama ok')
            driver.close()
        elif len(lista) == 0:
            WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.NAME, 'reportDate'))).send_keys(Keys.DOWN)
            driver.find_element_by_partial_link_text("Faktura").click()
            time.sleep(3)
            print('Proama ok')
            driver.close()
        else:
            print('brak sprzedaży w Proamie')
            driver.close()

            time.sleep(1)

    except:
        print('brak Proamy')
        driver.close()
        pass





### UNIQA ###
def uniqa():
    options = webdriver.ChromeOptions()
    preferences = {'download.default_directory': "C:\\Users\\ROBERT\\Desktop\\Księgowość\\2019"}
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
        print('Brak UNIQA')
        pass



### WARTA ###
def warta():
    # try:
        options = webdriver.ChromeOptions()
        preferences = {'download.default_directory' : "C:\\Users\\ROBERT\\Desktop\\Księgowość\\2019"}
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
            # time.sleep(2.1)
            # driver.find_element_by_css_selector().click()
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH,
                                                                           "//*[contains(text(), 'Majątek')]"))).click()
            time.sleep(0.9)
            rozliczenia_agencji = 'https://eagent.warta.pl/view360/#/app/main/settlement/property/LODD01643002/agent/list/?aid=1770160&agentOuid=LODD01643002'
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
            print('Brak WARTA')
            pass



### WARTA ŻYCIE ###
def warta_ż():
    options = webdriver.ChromeOptions()
    preferences = {'download.default_directory': "C:\\Users\\ROBERT\\Desktop\\Księgowość\\2019",
                   'plugins.always_open_pdf_externally': True}
    options.add_experimental_option("prefs", preferences)
    driver = webdriver.Chrome(executable_path=r'M:/zzzProjekty/drivery przegądarek/chromedriver.exe', options=options)

    try:
        url_warta_ż = 'https://eplatforma.warta.pl/'
        driver.get(url_warta_ż)

        WebDriverWait(driver, 4).until(EC.presence_of_element_located((By.NAME, "LOGNAME_13"))).send_keys(warta_ż_l)
        WebDriverWait(driver, 4).until(EC.presence_of_element_located((By.NAME, "PASSWD_13"))).send_keys(warta_ż_h)
        WebDriverWait(driver, 4).until(EC.presence_of_element_located((By.NAME, "zaloguj"))).click()

        WebDriverWait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR,
                                                                       "#Cont\.srodek\.1 > div > form > table > tbody > tr.filtr > th:nth-child(2) > select"))).click()

        WebDriverWait(driver, 4).until(
            EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'fakturowanie')]"))).click()
        WebDriverWait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR,
                                                                       "#Cont\.srodek\.1 > div > form > table > tbody > tr.filtr > th:nth-child(8) > input:nth-child(1)"))).click()
        WebDriverWait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR,
                                                                       "#Cont\.srodek\.1 > div > form > table > tbody > tr.td_line2 > td:nth-child(8) > a > img"))).click()
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
    preferences = {'download.default_directory': "C:\\Users\\ROBERT\\Desktop\\Księgowość\\2019"}
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
    preferences = {'download.default_directory': "C:\\Users\\ROBERT\\Desktop\\Księgowość\\2019"}
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




if __name__ == '__main__':
    # multiprocessing.freeze_support()

    p1 = multiprocessing.Process(target=santander)
    p2 = multiprocessing.Process(target=allianz)
    p3 = multiprocessing.Process(target=main)        ### AXA | WIENER | INSLY ###
    p4 = multiprocessing.Process(target=compensa)
    p5 = multiprocessing.Process(target=generali)
    p6 = multiprocessing.Process(target=hestia)
    p7 = multiprocessing.Process(target=interrisk)
    p8 = multiprocessing.Process(target=proama)
    p9 = multiprocessing.Process(target=uniqa)
    p10 = multiprocessing.Process(target=warta)
    p11 = multiprocessing.Process(target=warta_ż)
    p12 = multiprocessing.Process(target=unilink)
    p13 = multiprocessing.Process(target=pzu)

##################

    p1.start()
    p2.start()
    p3.start()

    p1.join()
    p2.join()
    p3.join()

###################

    p4.start()
    p5.start()
    p6.start()

    p4.join()
    p5.join()
    p6.join()

####################

    p7.start()
    p8.start()
    p9.start()

    p7.join()
    p8.join()
    p9.join()

####################

    p10.start()
    p11.start()

    p10.join()
    p11.join()

####################

    p12.start()
    p13.start()

    p12.join()
    p13.join()
