import os

import pandas as pd
import numpy as np
import string
import io


pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', 120)


df = pd.DataFrame({
     'int32':    [['ada', 'ala'], ['ula'], None, None],
     'int64':    np.random.randint(10**7, 10**9, 4).astype(np.int64)*10,
     'float':    np.random.rand(4),
     'string':   np.random.choice(['jed', 'dwa', None, None], 4),
     'string2':   np.random.choice(['jed', 'dwa', None, None], 4),
     })

# print(df)


df_pickle = pd.read_pickle('san.pkl')[:]
# df_pickle = pd.DataFrame(df_pickle)
print(df_pickle)


store = pd.HDFStore('example.h5')


# store.append('store_key16', df_pickle, format='fixed', append=False)

# print(store.info())
# store.close()

# df = pd.read_hdf(store, key='store_key16')
# store.close()
# print(df)










# df = pd.read_pickle('san.pkl')
# print(df['class'])
#
# hdf = pd.HDFStore('storage.h5')
# # hdf['class'] = df.astype(str)
# print(hdf.groups)
#
# df = pd.DataFrame(df['class'], columns=('I', 'Aas'))# put the dataset in the storage
# hdf.put('d1', df, format='table')

#
# df1 = hdf.get_storer('san1')
# print()
# print(df1)














# df = pd.read_csv('san.csv')
# df = pd.read_pickle('san.pkl')
# print(df)
# df = pd.read_hdf('storage.h5', key='san1')
# df = pd.read_csv('Test.csv', header=16, on_bad_lines='skip')

# df = pd.read_csv('Test.csv', header=0, usecols=lambda c: c in df.columns)
# df = df.replace('\n', '\u2063', regex=True)
# to_test = pd.read_csv('Test.csv', dtype=object,
#                       header=0)



# df1 = df.iloc[[2, 205]]
# df2 = df.iloc[[51, 53]]
# df3 = df.iloc[[53]]
# df4 = df.iloc[[37, 38]]
# df666 = df#.iloc[[139]]

# df1 = df1.replace('\u2063', '\n', regex=True)

# df = df.drop([2, 3], axis=0)
# print(df4)
# print(df666['class'])
#
# df.to_pickle('san.pkl', mod)

# df5.to_csv('Test.csv', mode='a', index=False, sep=',')
# df5.to_csv('Test.csv', mode='a', index=False, sep=',')
# df1.to_csv('Test.csv', index=False, sep=',')
# df3.to_csv('Test.csv', mode='a', index=False, sep=',')
# df666.to_csv('Test.csv', mode='a', index=False, sep=',')
