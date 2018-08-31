# coding=utf-8 #
# Author GJN #
import xlrd
import xlwings as xw
import pandas as pd
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import Column, String, Date, Integer, create_engine, ForeignKey, Float
import datetime


def rs(year):
    route = r'E:\Company\League\PFL\PFL_playersheet %s-%s.xlsm' % (year, year + 1)
    season = '%s-%s' % (year, year + 1)
    return route, season


engine = create_engine('mysql+mysqlconnector://root:password@localhost:3306/pfl')
sql_cmd = "SELECT * FROM team"
df = pd.read_sql(sql_cmd, engine)
path = r'E:\Company\League\PFL\\'
ndf = pd.read_excel(path + 'team name.xlsx')
# df2 = df[df['team'] == 'Belenenses']
dlist = []
for name in ndf['PS'].tolist():
    dlist.append(df[df['team'] == name])


wb = xw.Book()
st = wb.sheets[0]

df16 = df[df['info'].str.contains('16-17')]
avd = df16.groupby('info').mean()
avsup = (avd.iloc[1, 4] - avd.iloc[0, 4]) / 2
avttg = (avd.iloc[1, 5] + avd.iloc[0, 5]) / 2


def h_sub_a(dfin):
    df2 = dfin.set_index('team')
    dfx = df2[df2['info'].str.contains('Home')].iloc[:, 1:].sub(df2[df2['info'].str.contains('Away')].iloc[:, 1:])
    return dfx


def sup(dfin, dfha):
    df2 = dfin.set_index('team')
    dfx = df2[df2['info'].str.contains('Home')]
    dfx['S'] = dfx['sup'] - dfha['HA']
    dfx['G'] = dfx['gd'] - dfha['HA']
    return dfx


def ttg(dfin):
    df2 = dfin.set_index('team')
    dfx = df2[df2['info'].str.contains('Home')].iloc[:, 1:].add(df2[df2['info'].str.contains('Away')].iloc[:, 1:])
    dfx2 = dfx / 2
    return dfx2


df16x = h_sub_a(df16)
df16x['HA'] = df16x['sup'] - avsup
df16s = sup(df16, df16x)

df17 = df[df['info'].str.contains('17-18')]
avd17 = df16.groupby('info').mean()
avsup17 = (avd17.iloc[1, 4] - avd17.iloc[0, 4]) / 2
avttg17 = (avd17.iloc[1, 5] + avd17.iloc[0, 5]) / 2
df17x = h_sub_a(df17)
df17x['HA'] = df17x['sup'] - avsup17
df17s = sup(df17, df17x)

df15 = df[df['info'].str.contains('15-16')]
avd15 = df16.groupby('info').mean()
avsup15 = (avd15.iloc[1, 4] - avd15.iloc[0, 4]) / 2
avttg15 = (avd15.iloc[1, 5] + avd15.iloc[0, 5]) / 2
df15x = h_sub_a(df15)
df15x['HA'] = df15x['sup'] - avsup15
df15s = sup(df15, df15x)

df14 = df[df['info'].str.contains('14-15')]
avd14 = df16.groupby('info').mean()
avsup14 = (avd14.iloc[1, 4] - avd14.iloc[0, 4]) / 2
avttg14 = (avd14.iloc[1, 5] + avd14.iloc[0, 5]) / 2
df14x = h_sub_a(df14)
df14x['HA'] = df14x['sup'] - avsup14
df14s = sup(df14, df14x)
i = 1
for df1 in [df17s, df16s, df15s, df14s]:
    st.range(1, i).value = df1[['S', 'G']]
    i += 4
i = 1
for df1 in [df17x, df16x, df15x, df14x]:
    st.range(1, i + 2).value = df1[['HA']]
    i += 4

df17t = ttg(df17)
df16t = ttg(df16)
df15t = ttg(df15)
df14t = ttg(df14)

i = 1
for df1 in [df17t, df16t, df15t, df14t]:
    st.range(1, i).value = df1.iloc[:, [3, 5]]
    i += 4

