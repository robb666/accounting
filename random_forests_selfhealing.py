import os
import time

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
    tags = ['a', 'input', 'div']  # to uzupełniać..?
    # el_name = ['LOGIN', 'PASSW', 'LOG_BUTTON']
    arr = []
    for tag in tags:
        for element in soup.find_all(tag):
            arr.append({'tag': tag} |
                       {'text': element.text} |
                       element.attrs)

    df = pd.DataFrame.from_records(arr)
    df.insert(0, 'element', df.title)
    df.element.fillna(df['id'], inplace=True)
    df.element.fillna(df['text'], inplace=True)
    # df.text = np.nan  # fillna didn't work
    # df = df.replace('\n', '\u2063', regex=True)

    # print(df)
    ########

    # df.to_csv('san.csv', index=False, sep=',', encoding='utf-8')
    # df.to_pickle('san.pkl')
    df = pd.DataFrame(dtype=object).convert_dtypes()
    hdf = pd.HDFStore('storage.h5')
    hdf.put('san1', df, format='table', data_columns=True)
    hdf.close()
    # # df = pd.read_csv('san.csv', dtype=object, converters={'some_name':lambda x:x.replace('/n','')})
    # df = pd.read_csv('san.csv', dtype=object)

    return df


def healed_locator(driver, *, helper_attr, header, element_row, value, filename='Test.csv'):
    df = scrp(driver)
    df = df.fillna('None')
    df = df.replace('\u2063', '\n', regex=True)
    print()
    to_test = pd.read_csv(filename, dtype=object, header=header,
                          usecols=lambda c: c in df.columns).iloc[[element_row]]

    to_test = to_test.fillna('None')
    to_test = to_test.replace('\u2063', '\n', regex=True)

    processed_test = pd.concat([df, to_test], axis=0)
    print(processed_test)

    processed_test = processed_test.iloc[[-1]]

    ohe = OneHotEncoder(sparse=False, handle_unknown='ignore')
    X_train = ohe.fit_transform(df.astype(str))
    X_test = ohe.transform(processed_test)

    element_dict = dict(zip(df['element'].unique(), range(df['element'].nunique())))
    y_train = df['element'].replace(element_dict)

    rf = RandomForestClassifier(n_estimators=50, random_state=0)
    rf.fit(X_train, y_train)

    probabilities = rf.predict_proba(X_test)[0]
    print(probabilities)
    el_attr = list(element_dict.keys())[np.argmax(probabilities)]

    columns = df.columns[df.isin([el_attr]).any()].values  # kolumny atrybutu
    # TODO zakwalifikować atrybut..bez iteracji
    for attr in columns[1:]:
        print(f"//*[@{attr}='{el_attr}' {helper_attr}]")
        try:  # kiedy więcej niż jeden element o danym atrybucie znajduje się na stronie.
            selector = driver.find_element(By.XPATH, f"//*[@{attr}='{el_attr}' {helper_attr}]")
            if value:
                selector.send_keys(value)
            else:  # click
                WebDriverWait(driver, 4).until(EC.element_to_be_clickable((
                    By.XPATH, f"//*[@{attr}='{el_attr}' {helper_attr}]"))).click()
            break  # return
        except Exception as e:
            print(e)
