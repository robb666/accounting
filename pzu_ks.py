from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from L_H_ks import pzu_h, pzu_l
import time




### PZU ###
def pzu():
    ### CHROME ###
    options = webdriver.ChromeOptions()
    preferences = {'download.default_directory': "C:\\Users\\ROBERT\\Desktop\\Księgowość\\2019"}
    options.add_experimental_option("prefs", preferences)
    driver = webdriver.Chrome()#executable_path=r'M:/zzzProjekty/drivery przegądarek/chromedriver.exe', options=options)

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
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, 'ProducerStatementReportOnlinePzu:0:StatementsLV:0:DownloadPdfFileLink'))).click()
        time.sleep(2)
        driver.quit()
        print('PZU ok')

    except:
        driver.quit()
        print('Brak PZU')
        pass


pzu()