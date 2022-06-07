import pandas as pd
import os
import sys
import subprocess

# os.environ['PATH'] += \
#     os.pathsep + os.path.expanduser(r'C:\Users\PipBoy3000\Desktop\IT\projekty\accounting\.env\Lib\site-packages\tables.libs')



store = pd.HDFStore(os.getcwd() + '\elements.h5')
# print(store.info())


class Elements:
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


os.getcwd()
el = Elements()
print(el.accept_1)