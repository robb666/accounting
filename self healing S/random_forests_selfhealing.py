import pandas as pd
import numpy as np
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from sklearn.preprocessing import OneHotEncoder
from sklearn.ensemble import RandomForestClassifier


pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)


def scrp(driver):
    html = driver.page_source
    soup = BeautifulSoup(html, 'lxml')
    tags = ['a']  # to uzupełniać..?
    # el_name = ['LOGIN', 'PASSW', 'LOG_BUTTON']
    arr = []
    for tag in tags:
        for element in soup.find_all(tag):
            arr.append({'tag': tag} |
                       {'text': element.text} |
                       element.attrs)

    df = pd.DataFrame.from_records(arr)
    df.insert(0, 'element', df.title)
    df.element.fillna(df['text'], inplace=True)
    # df.text = np.nan  # fillna didn't work
    # df = df.replace('\n', '\u2063', regex=True)
    # print(df)
    ########
    # df.to_csv('san.csv', index=False, sep=',', encoding='utf-8')
    # # df = pd.read_csv('san.csv', dtype=object, converters={'some_name':lambda x:x.replace('/n','')})
    # df = pd.read_csv('san.csv', dtype=object)

    return df


def healed_locator(driver, e, *, attr, element_row, value):
    if 'no such element' in str(e) or 'Unable to locate element' in str(e) or 'element not interactable' in str(e):

        df = scrp(driver)
        # df = df.replace('\u2063', '\n', regex=True)
        df = df.fillna('None')
        # df = df.drop(['Unnamed: 0'], axis=1)
        print(df.head(15))
        # df = df.head()

        # to_test = pd.read_csv('Test.csv').iloc[[element_row]]
        to_test = pd.read_csv('Test.csv', dtype=object,
                              header=0, usecols=lambda c: c in df.columns)
        to_test = to_test.replace('\u2063', '\n', regex=True)
        print('to_test')
        print(to_test)
        # processed_test = to_test.append(df)[df.columns].iloc[[element_row]]

        processed_test = pd.concat([df, to_test], axis=0)
        processed_test = processed_test.iloc[[-1]]

        print('processed_test')
        print(processed_test)




        # test = to_test.fillna('None')
        # concatenated = pd.concat([df, test], axis=0)#.drop('element', axis=1)
        # print(concatenated)

        # processed_test = pd.DataFrame(concatenated.iloc[-1]).T
        # processed_test = processed_test.drop(['Unnamed: 0'], axis=1)

        # print(processed_test)

        ohe = OneHotEncoder(sparse=False, handle_unknown='ignore')
        X_train = ohe.fit_transform(df.astype(str))
        X_test = ohe.transform(processed_test)

        element_dict = dict(zip(df['element'].unique(), range(df['element'].nunique())))
        y_train = df['element'].replace(element_dict)

        rf = RandomForestClassifier(n_estimators=50, random_state=0)
        rf.fit(X_train, y_train)

        probabilities = rf.predict_proba(X_test)[0]
        print(probabilities)
        print(len(probabilities))

        print(element_dict)
        el_attr = list(element_dict.keys())[np.argmax(probabilities)]
        print(f"//*[@{attr}='{el_attr}']")
        selector = driver.find_element(By.XPATH, f"//*[@{attr}='{el_attr}' and @class='button primary small']")
        print(selector)
        if value:
            selector.send_keys(value)
        else:  # click
            selector.click()
