import pandas as pd
import numpy as np
import os
import sys
import h5py
import subprocess
# store = pd.HDFStore(os.getcwd() + '\elements.h5')



# # store = h5py.File('elements.h5', 'r')
with h5py.File('elements.h5') as hdf:
    ls = list(hdf.keys())
    print(f'List of datasets: {ls}')
    data = list(hdf.items())
    print(list(data[0][1].get('axis0')[:]))
    # data = hdf.get('nik')
    # print(list(data['axis0']))
    # print(list(data['axis1']))
    # print(list(data['block0_items']))
    # print(list(data['block0_values']))
    # print(list(data))

    # print(list(hdf))
    # print(list(hdf.get('nik')))
    # base_items = list(hdf.items())
    # print(base_items)
    # G2 = hdf.get('ordinarypin')
    # G2_items = list(G2.items())
    # print(G2_items)
    # G21 = G2.get('axis0')
    # dataset = list(G21)
    #
    # print(dataset)
    #
    # df = pd.DataFrame(dataset).T
    # print(df)
    # Gn = G2.get('zaloguj')
    # print(Gn)

    # print(Gn)
# for i in store:
#     print(store[i])




print()
print()
print()
print()





store = pd.HDFStore(os.getcwd() + '\elements.h5')
print(store.info())
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
print(el.nik)
