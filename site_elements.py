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


class HDF5:

    def __init__(self, from_pickle, to_store, *, row=None):
        self.pickle = from_pickle
        self.store = pd.HDFStore(to_store)
        self.row = row

    def info(self):
        return self.store.info()

    def read_pickle(self):
        df = pd.read_pickle(self.pickle)
        return df.iloc[0:len(df), :] if self.row is None else df.iloc[self.row:self.row+1, :]

    def read(self, key):
        return pd.read_hdf(self.store, key=key)

    def append(self, key):
        return self.store.append(key, self.read_pickle(), format='fixed', append=False)

    def remove(self, element):
        return self.store.remove(element)

    def close(self):
        return self.store.close()


pickle_file = r'san.pkl'
store_file = r'elements.h5'

# hdf = HDF5(pickle_file, store_file, row=50)
# print(hdf.read_pickle())
# hdf.append('myśl')
# hdf.remove('słowniczek')

# print(hdf.read('myśl'))
# print(hdf.info())


#
# hdf.close()
