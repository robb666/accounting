import pandas as pd
import numpy as np
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from sklearn.preprocessing import OneHotEncoder
from sklearn.ensemble import RandomForestClassifier

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', 120)


def scrp(driver):
    html = driver.page_source
    soup = BeautifulSoup(html, 'lxml')
    tags = ['a', 'input', 'div', 'span']  # to uzupełniać..
    arr = []
    for tag in tags:
        for element in soup.find_all(tag):
            arr.append({'tag': tag} |
                       {'text': element.text} |
                       element.attrs)
    # print(arr)
    df = pd.DataFrame.from_records(arr)
    df.insert(0, 'element', df.title)
    df.element.fillna(df['id'], inplace=True)
    df.element.fillna(df['text'], inplace=True)

    # df.to_pickle('san.pkl')

    return df



def healed_locator(driver, *, helper_attr='', element, value):
    df = scrp(driver)
    df = df.fillna('None')

    to_test = pd.DataFrame(element, dtype=object, columns=df.columns)
    to_test = to_test.fillna('None')

    processed_test = pd.concat([df, to_test], axis=0)
    # print(processed_test)
    processed_test = processed_test.iloc[[-1]]

    ohe = OneHotEncoder(sparse=False, handle_unknown='ignore')
    X_train = ohe.fit_transform(df.astype(str))
    X_test = ohe.transform(processed_test.astype(str)) ###str???

    element_dict = dict(zip(df['element'].unique(), range(df['element'].nunique())))
    y_train = df['element'].replace(element_dict)

    rf = RandomForestClassifier(n_estimators=50, random_state=0)
    rf.fit(X_train, y_train)

    probabilities = rf.predict_proba(X_test)[0]
    print(probabilities)

    el_attr = list(element_dict.keys())[np.argmax(probabilities)]
    # print('el_attr')
    # print(el_attr)
    columns = df.columns[df.isin([el_attr]).any()].values  # kolumny atrybutu
    # TODO zakwalifikować atrybut..bez iteracji, wykożystać np.isin, ----||---- indeks ze sklepu, Page Object Model (POM)

    for attr in columns[1:]:

        try:  # kiedy więcej niż jeden element o danym atrybucie znajduje się na stronie.

            if attr == 'text':
                selector = driver.find_element(By.XPATH, f"//*[contains({attr}(), '{el_attr}') {helper_attr}]")
                print(f"//*[contains({attr}(), '{el_attr}') {helper_attr}]")
            else:
                selector = driver.find_element(By.XPATH, f"//*[@{attr}='{el_attr}' {helper_attr}]")
                print(f"//*[@{attr}='{el_attr}' {helper_attr}]")

            if value:
                selector.send_keys(value)
            else:  # click
                WebDriverWait(driver, 4).until(EC.element_to_be_clickable((
                                                        By.XPATH, f"//*[@{attr}='{el_attr}' {helper_attr}]"))).click()
            break  # return
        except Exception as e:
            print(e)
