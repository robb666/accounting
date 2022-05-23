import pandas as pd


pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
# pd.set_option('display.max_colwidth', None)

df = pd.read_csv('san.csv')
# df = pd.read_csv('Test.csv', header=6, on_bad_lines='skip')

# df = pd.read_csv('Test.csv', header=0, usecols=lambda c: c in df.columns)
# df = df.replace('\n', '\u2063', regex=True)
# to_test = pd.read_csv('Test.csv', dtype=object,
#                       header=0)



# df1 = df.iloc[[2, 205]]
# df2 = df.iloc[[51, 53]]
# df3 = df.iloc[[53]]
# df4 = df.iloc[[37, 38]]
df5 = df.iloc[[34, 35]]

# df1 = df1.replace('\u2063', '\n', regex=True)

# df = df.drop([2, 3], axis=0)
# print(df4)
print(df5)

# df5.to_csv('Test.csv', mode='a', index=False, sep=',')
# df1.to_csv('Test.csv', index=False, sep=',')
# df3.to_csv('Test.csv', mode='a', index=False, sep=',')
