import pandas as pd
import os


pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', 120)


class Elements:
    store = pd.HDFStore(os.getcwd() + r'\elements.h5')

    accept_1 = pd.read_hdf(store, key='accept_1')
    login = pd.read_hdf(store, key='zaloguj')
    san = pd.read_hdf(store, key='SInternet')
    accept_2 = pd.read_hdf(store, key='accept_2')
    nik = pd.read_hdf(store, key='nik')
    button_nik = pd.read_hdf(store, key='button_nik')
    ordinarypin = pd.read_hdf(store, key='ordinarypin')
    button_ordinarypin = pd.read_hdf(store, key='button_ordinarypin')
    oneTimeAccess = pd.read_hdf(store, key='oneTimeAccess')
    otp = pd.read_hdf(store, key='otp')
    otp_button = pd.read_hdf(store, key='otp_button')

    store.close()


class HDF:
    def __init__(self, pickle, store, element=None):
        self.pickle = pickle
        self.store = store
        self.element = element

    def store(self):
        return pd.HDFStore('elements.h5')

    def read_pickle(self):
        return pd.read_pickle(self.pickle).iloc[[11]]

    def append(self):
        return self.store.append('otp_button', self.read_pickle, format='fixed', append=False)

    def remove(self):
        return self.store.remove(self.element)

    def read(self):
        return pd.read_hdf(self.store, key='nik')

    def info(self):
        return self.store.info()

    def __repr__(self):
        return self.read_hdf

    def close(self):
        return self.store.close()


file = r'san.pkl'
store = r'\elements.h5'
hdf = HDF(file, store)

print(hdf.read_pickle())
# print(hdf.info())
hdf.close_hdf()
