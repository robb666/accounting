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
import concurrent.futures
from faktury_GmailAPI import email, zallianz
from cash_excel import raport_inkaso
from L_H_ks import san_l, san_h, allianz_l, allianz_h, compensa_l, compensa_h, eins_l, eins_h, generali_l, generali_h, \
     hestia_l, hestia_h, uniqa_l, uniqa_h, warta_l, warta_h, interrisk_l, interrisk_h, proama_l, proama_h, \
     unilink_l, unilink_h, pzu_l, pzu_h, warta_ż_l, warta_ż_h, gapi, bookkeeping, OTP
import time
import smtplib, ssl
from email import encoders
import mimetypes
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from dateutil.relativedelta import relativedelta
from functools import wraps
import pyotp


def driver_inst(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        options = webdriver.ChromeOptions()
        preferences = {'download.default_directory': next_month_path,
                       'plugins.always_open_pdf_externally': True}
        options.add_experimental_option("prefs", preferences)
        driver = webdriver.Chrome(executable_path=r'M:/zzzProjekty/drivery przegądarek/chromedriver.exe',
                                  options=options)
        return func(driver)
    return wrapper


@driver_inst
def allianz(driver, url_allianz='https://start.allianz.pl'):
    try:
        for _ in range(2):
            driver.get(url_allianz)
            login = driver.find_element_by_id('username')
            login.send_keys(allianz_l)
            pwd = driver.find_element_by_id('password')
            pwd.send_keys(allianz_h)
            driver.find_element_by_name('submit').click()
            WebDriverWait(driver, 15).until(EC.url_changes(url_allianz))

        time.sleep(5.5)
        token = driver.find_element_by_id('token')
        tiktok = zallianz()
        token.send_keys(tiktok)
        driver.find_element_by_xpath('//button[@accesskey="s"]').click()
        url_inv = 'https://chuck.allianz.pl/agent/#/invoices'
        driver.get(url_inv)
        WebDriverWait(driver, 15).until(EC.url_changes(url_inv))
        WebDriverWait(driver, 9).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.am-btn-large'))).click()
        time.sleep(1)
        WebDriverWait(driver, 9).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.am-btn-primary')))[1].click()
        time.sleep(4)
        driver.quit()
        print('Allianz ok')
    except:
        with open(rf"{next_month_path}/brak dokumentów.txt", "a") as f:
            f.write("Brak Allianz\n")
        print('Brak Allianz')
        driver.quit()


@driver_inst
def compensa(driver, url_compensa='https://cportal.compensa.pl/'):
    try:
        driver.get(url_compensa)
        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.input'))).send_keys(compensa_l)
        driver.find_element_by_css_selector('div.fl:nth-child(5) > input:nth-child(1)').send_keys(compensa_h)
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "btnLogin"))).click()
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, 'News')))
        url_compensa = 'https://cportal.compensa.pl/#MyCommissions'
        time.sleep(1)
        driver.get(url_compensa)
        driver.set_page_load_timeout(30)
        time.sleep(3)
        try:
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH,
                                                                           "//*[contains(text(), 'Zamkn')]"))).click()
        except:
            pass
        time.sleep(1)
        driver.get_screenshot_as_file(f'{next_month_path}/Compensa.png')
        driver.close()
        print('Compensa ok')
    except Exception as e:
        driver.close()
        with open(rf"{next_month_path}/brak dokumentów.txt", "a") as f:
            f.write("Brak Compensa\n")
        print('Brak Compensa')


@driver_inst
def euroins(driver, url_eins='https://eins.com.pl/index.php/login'):
    try:
        driver.get(url_eins)
        driver.find_element_by_xpath('//input[@id="user"]').send_keys(eins_l)
        driver.find_element_by_xpath('//input[@id="password"]').send_keys(eins_h)
        driver.find_element_by_xpath('//input[@id="submit-form"]').click()
        totp = pyotp.TOTP(OTP).now()
        WebDriverWait(driver, 9).until(EC.presence_of_element_located((By.XPATH,
                                                                '//input[@name="challenge"]'))).send_keys(totp)
        driver.find_element_by_xpath('//button[@type="submit"]').click()
        WebDriverWait(driver, 9).until(EC.url_changes(url_eins))
        url = 'https://eins.com.pl/index.php/apps/files/?dir=/MAGRO_E/noty%20prowizyjne&fileid=58566'
        driver.get(url)
        WebDriverWait(driver, 9).until(EC.url_contains('58566'))
        driver.find_element_by_xpath(
            '//tbody/tr[not(@data-id <= preceding-sibling::tr/@data-id) and not(@data-id <= following-sibling::tr/@data-id)]'
        ).click()
        time.sleep(5)
        driver.quit()
        print('Euroins ok')
    except:
        time.sleep(1)
        with open(rf"{next_month_path}/brak dokumentów.txt", "a") as f:
            f.write("Brak Euroins\n")
        print('Brak Euroins')
        driver.quit()


@driver_inst
def generali(driver,
             url_generali='https://portal.generali.pl/auth/login?service=https%3A%2F%2Fportal.generali.pl%2Flogin%2Fcas'):
    try:
        driver.get(url_generali)
        WebDriverWait(driver, 9).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#username'))).send_keys(generali_l)
        driver.find_element_by_css_selector('#password').send_keys(generali_h)
        driver.find_element_by_css_selector('#fm1 > div.login > input[type=submit]:nth-child(6)').click()
        url_accounting = 'https://portal.generali.pl/mikado/commissions/current'
        driver.get(url_accounting)
        WebDriverWait(driver, 9).until(EC.presence_of_element_located((By.XPATH, "//*[@class='far fa-file-zip-o']"))).click()
        time.sleep(10)
        driver.close()
        print('Generali ok')
    except:
        driver.close()
        with open(rf"{next_month_path}/brak dokumentów.txt", "a") as f:
            f.write("Brak Generali\n")
        print('Brak Generali')


@driver_inst
def hestia(driver, url='https://sso.ergohestia.pl/my.policy'):
    try:
        driver.get(url)
        login_hes = driver.find_element_by_id('input_1').send_keys(hestia_l)
        hasło_hes = driver.find_element_by_id('input_2').send_keys(hestia_h)
        WebDriverWait(driver, 9).until(EC.presence_of_element_located((By.XPATH,
                                                                '//*[@id="auth_form"]/div[3]/div[4]/button'))).click()
        url_agent = 'https://partner.ergohestia.pl/#/partner'
        driver.get(url_agent)
        url_agent = 'https://partner.ergohestia.pl/#/partner' + '/commissionHistory'
        driver.get(url_agent)
        WebDriverWait(driver, 11).until(EC.presence_of_element_located((By.XPATH,
                                                                        '/html/body/div/div[3]/div[2]/div/ng-include/'
                                                                        'div/div[3]/div/div/div/div[2]/table/tbody/'
                                                                        'tr[1]/td[6]/ng-include/a'))).click()
        time.sleep(2.2)
        faktura = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,
                                                                        '/html/body/div/div[3]/div[2]/div/ul/li[2]/a')))
        faktura.click()
        time.sleep(5)
        driver.quit()
        print('Hestia ok')
    except:
        driver.quit()
        with open(rf"{next_month_path}\brak dokumentów.txt", "a") as f:
            f.write("Brak Hestii\n")
        print('Brak Hestii')


@driver_inst
def interrisk(driver, url_interrisk='https://portal.interrisk.pl/Zaloguj'):
    try :
        driver.get(url_interrisk)
        driver.find_element_by_id('ctl00_cph1_uxLogin_UserName').send_keys(interrisk_l)
        WebDriverWait(driver, 9).until(EC.presence_of_element_located((By.ID,
                                                                "ctl00_cph1_uxLogin_Password"))).send_keys(interrisk_h)
        WebDriverWait(driver, 9).until(EC.presence_of_element_located((By.ID, "ctl00_cph1_uxLogin_LoginButton"))).click()

        url_interrisk_prow = 'https://portal.interrisk.pl/Rozliczenia/NotyProwizyjne/Przegladaj'
        driver.get(url_interrisk_prow)
        WebDriverWait(driver, 9).until(EC.presence_of_element_located((By.CSS_SELECTOR,
                                                                       "#ctl00_ctl00_cph1_cph1_search"))).click()
        WebDriverWait(driver, 9).until(EC.presence_of_element_located((By.XPATH,
                            '/html/body/form/div[3]/div[2]/div[1]/div[2]/div[2]/table/tbody/tr[1]/td[6]/input'))).click()
        WebDriverWait(driver, 9).until(EC.presence_of_element_located((By.ID,
                                                                        'ctl00_ctl00_cph1_cph1_cbNoteOnDemand'))).click()
        time.sleep(1.5)
        WebDriverWait(driver, 9).until(EC.presence_of_element_located((By.CSS_SELECTOR,
                                                                        "#ctl00_ctl00_cph1_cph1_exportPdf"))).click()
        WebDriverWait(driver, 9).until(EC.presence_of_element_located((By.CSS_SELECTOR,
                                                                       "#ctl00_ctl00_cph1_cph1_search"))).click()
        try :
            WebDriverWait(driver, 9).until(EC.presence_of_element_located((By.XPATH,
                            '/html/body/form/div[3]/div[2]/div[1]/div[2]/div[2]/table/tbody/tr[2]/td[6]/input'))).click()
            WebDriverWait(driver, 9).until(EC.presence_of_element_located((By.CSS_SELECTOR,
                                                                            "#ctl00_ctl00_cph1_cph1_exportPdf"))).click()
        except :
            pass
        time.sleep(2.5)
        driver.quit()
        print('InterRisk ok')
    except :
        driver.quit()
        with open(rf"{next_month_path}/brak dokumentów.txt", "a") as f:
            f.write("Brak InterRisk\n")
        print('Brak InterRisk')


@driver_inst
def uniqa(driver, url_post_login='https://pos.uniqa.pl/pl/login_fe'):
    payload = {'login': uniqa_l,
               'password': uniqa_h}
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                                                                                    'Chrome/73.0.3683.86 Safari/537.36'}
    try:
        with requests.Session() as s:
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
        with open(rf"{next_month_path}/brak dokumentów.txt", "a") as f:
            f.write("Brak UNIQA\n")
        print('Brak UNIQA')


@driver_inst
def warta(driver,
          url_warta='https://cas.warta.pl/cas/login?service=https%3A%2F%2Feagent.warta.pl%2Fview360%2Flogin%2Fcas'):
    try :
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
        rozliczenia_agencji = 'https://eagent.warta.pl/view360/#/app/main/settlement/property/A00005152001/agent/list' \
                              '/?aid=1770143&agentOuid=A00005152001'
        driver.get(rozliczenia_agencji)
        time.sleep(1.1)
        WebDriverWait(driver, 4).until(EC.presence_of_all_elements_located((By.XPATH,
                                                                            "//*[contains(text(), 'RSP')]")))[0].click()
        WebDriverWait(driver, 4).until(EC.presence_of_all_elements_located((By.CLASS_NAME,
                                                "settlement-details-documents__content__item__list__elem")))[2].click()
        driver.get(rozliczenia_agencji)
        time.sleep(0.9)
        WebDriverWait(driver, 4).until(EC.presence_of_all_elements_located((By.XPATH,
                                                                            "//*[contains(text(), 'RSP')]")))[1].click()
        WebDriverWait(driver, 4).until(EC.presence_of_all_elements_located((By.CLASS_NAME,
                                                "settlement-details-documents__content__item__list__elem")))[2].click()
        time.sleep(2)
        driver.quit()
        print('Warta ok')
    except:
        driver.quit()
        with open(rf"{next_month_path}/brak dokumentów.txt", "a") as f:
            f.write("Brak WARTA\n")
        print('Brak WARTA')
        pass


@driver_inst
def warta_ż(driver, url_warta_ż='https://eplatforma.warta.pl/'):
    try:
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
        with open(rf"{next_month_path}/brak dokumentów.txt", "a") as f:
            f.write("Brak Warta Życie\n")
        print('Brak Warta Życie')


@driver_inst
def unilink(driver, url_unilink='https://unilink.pl/logowanie'):
    try:
        driver.get(url_unilink)
        WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.ID, "login"))).send_keys(unilink_l)
        WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.ID, "password"))).send_keys(unilink_h)
        WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.ID, "submit"))).click()
        time.sleep(1)
        url_unilink_faktury = 'https://unilink.pl/pokaz/4020'
        driver.get(url_unilink_faktury)
        try:
            WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.XPATH,
                                                '//*[@id="searchformlist"]/table/tbody/tr[2]/td[13]/div[1]/a/i'))).click()
        except:
            unilink()
        time.sleep(4)
        driver.quit()
        print('Unilink ok')
    except:
        driver.quit()
        with open(rf"{next_month_path}/brak dokumentów.txt", "a") as f:
            f.write("Brak Unilink\n")
        print('Brak Unilink')


@driver_inst
def pzu(driver, url_pezu='https://everest.pzu.pl/pc/PolicyCenter.do'):
    try:
        driver.get(url_pezu)
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
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR,
                                        '#Desktop\:MenuLinks\:Desktop_ProducerStatementReportOnlinePzu > div'))).click()
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID,
                                                'ProducerStatementReportOnlinePzu:0:statementTab-btnInnerEl'))).click()
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR,
                                '#ProducerStatementReportOnlinePzu\:0\:StatementsLV\:1\:DownloadPdfFileLink'))).click()
        time.sleep(4)
        driver.quit()
        documents = next_month_path
        for item in os.listdir(documents):
            if '+' in item and os.path.isfile(os.path.join(documents, item)):
                os.rename(os.path.join(documents, item), os.path.join(documents, 'PZU.pdf'))
        print('PZU ok')
    except:
        driver.quit()
        with open(rf"{next_month_path}brak dokumentów.txt", "a") as f:
            f.write("Brak PZU\n")
        print('Brak PZU')
        pass


def path_exists(next_month_path, num):
    if os.path.exists(next_month_path):
        num += 1
        next_month_path = f'{next_month_path[:-1]}..{str(num)}\\'
        return path_exists(next_month_path, num)
    else:
        return next_month_path


def mk_month_dir(next_month_dir):
    os.mkdir(next_month_dir)


def send_attachments(sender_email, receiver_email):
    msc_rok = (datetime.today() + relativedelta(months=-1)).strftime('%m.%Y')
    message = MIMEMultipart()
    message['Subject'] = f'Dokumenty za {msc_rok}'
    body = """Cześć, przesyłam dokumenty w załącznikach.\n\n"""
    message.attach(MIMEText(body))

    # TODO -> Dostęp dla mniej bezpiecznych aplikacji../zmiana na API Gmail | --> dodana wysyłka pustego maila w połowie okresu
    documents = next_month_path
    os.chdir(documents)
    for attachment in os.listdir(documents):
        content_type, encoding = mimetypes.guess_type(attachment, strict=False)
        if content_type is not None:
            main_type, sub_type = content_type.split('/', 1)
            my_file = MIMEBase(main_type, sub_type)
        else:
            pass

        with open(attachment, 'rb') as f:
            my_file.set_payload(f.read())
            my_file.add_header('Content-Disposition', f'attachment; filename = {attachment}', )
            encoders.encode_base64(my_file)
            message.attach(my_file)
            text = message.as_string()

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as server:
        server.login('ubezpieczenia.magro@gmail.com', gapi)
        server.sendmail(sender_email, receiver_email, text)


if __name__ == '__main__':
    # os.chdir(r'C:\Users\ROBERT\Desktop\IT\PYTHON\PYTHON 37 PROJEKTY\księgowość\skrypty osobno\dist')
    next_month_path = f'C:\\Users\\ROBERT\\Desktop\\Księgowość\\' \
                      f'{(datetime.today() + relativedelta(months=-1)).strftime("%m.%Y")}\\'

    next_month_path = path_exists(next_month_path, 0)
    mk_month_dir(next_month_path)

    tasks = [allianz, compensa, euroins, generali, hestia, interrisk, uniqa, warta, warta_ż, unilink, pzu]

    raport_inkaso(za_okres=-1, path=next_month_path)
    email(next_month_path)  # faktury z gmailAPI
    with concurrent.futures.ThreadPoolExecutor(max_workers=3) as executor:
        for n in range(len(tasks)):
            executor.submit(tasks[n])

    # send_attachments('ubezpieczenia.magro@gmail.com', bookkeeping) *
    time.sleep(1)
