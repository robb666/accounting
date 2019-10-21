from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys  import Keys
from selenium.webdriver.firefox.options import Options
import time
from L_H_ks import proama_l, proama_h


# for i in range(10):
### PROAMA ###
def proama():

    ### CHROME ###
    options = webdriver.ChromeOptions()
    preferences = {'download.default_directory': "C:\\Users\\ROBERT\\Desktop\\Księgowość\\2019"}
    options.add_experimental_option("prefs", preferences)
    driver = webdriver.Chrome()#executable_path=r'M:/zzzProjekty/drivery przegądarek/chromedriver.exe', options=options)

# try:

    url_proama = 'https://proagent.proama.pl/'
    driver.get(url_proama)
    WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, 'ZALOGUJ'))).click()
    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#username'))).send_keys(proama_l)
    driver.find_element_by_css_selector('#password').send_keys(proama_h)
    driver.find_element_by_css_selector('.login > input:nth-child(4)').click()
    url_accounting = 'https://portal.proama.pl/pipp/chooseCommission.do?'
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

# except:
    print('Brak Generali')
    driver.close()
    pass

proama()

