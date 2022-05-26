import pandas as pd


pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', 120)


df_pickle = pd.read_pickle('san.pkl')#.iloc[[32]]
print(df_pickle)


store = pd.HDFStore('elements.h5')

# store.append('otp_button', df_pickle, format='fixed', append=False)

# store.remove('nik')
print()
print(store.info())
# store.close()


df = pd.read_hdf(store, key='otp_button')

print(df)
store.close()

