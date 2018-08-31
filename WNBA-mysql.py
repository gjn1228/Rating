# coding=utf-8 #
# Author GJN #
import pandas as pd
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import Column, String, Date, Integer, create_engine, ForeignKey, Float
import mysql.connector

data = pd.read_excel(r'Y:\Departments\Risk Management\Level C\TeamSheet\WNBA\WNBA Teamsheet V2.xlsm', sheet_name='Results')

df = data.dropna(axis=0, how='all')
df1 = df.iloc[5:]
df1.columns = range(108)
df2 = df1.reset_index()[[0, 1, 2, 3, 7, 13, 14, 16, 45, 46, 47, 48, 49, 50, 53, 54, 55, 56, 57, 58]].reset_index()
df2.columns = ['index', 'season', 'league', 'stage', 'us_time', 'home', 'H', 'A', 'away',
               'sup_mkt_open', 'sup_mkt_close', 'act_sup', 'sup_dif', 'sup_csl_open', 'sup_csl_close',
               'hilo_mkt_open', 'hilo_mkt_close', 'act_hilo', 'hilo_dif', 'hilo_csl_open', 'hilo_csl_close']
Base = declarative_base()
# engine = create_engine('mysql+mysqlconnector://gjn:pass@172.18.1.158:3306/betradar')
engine = create_engine('mysql+mysqlconnector://root:password@localhost:3306/wnba')
dtype = {}
for c in df2.columns:
    i = list(df2.columns).index(c)
    if i in [0, 6, 7, 12, 18]:
        dtype[c] = Integer()
    elif i > 8:
        dtype[c] = Float()
    elif i == 4:
        dtype[c] = Date()
    else:
        dtype[c] = String(25)
df3 = df2.dropna(axis=0, how='any').reset_index()[df2.columns[1:]].reset_index()
df4 = df3[df3['season'] == '18'].reset_index()[df2.columns[1:]].reset_index()
df4.to_sql('2018', engine, schema='wnba', if_exists='replace', index=False, dtype=dtype)

# conn = mysql.connector.connect(user='gjn', password='pass',host='172.18.1.158', database='betradar', use_unicode=True)
# conn = mysql.connector.connect(user='root', password='password', database='wnba', use_unicode=True)
# cursor = conn.cursor()
# cursor.execute('ALTER TABLE teamsheet Add PRIMARY KEY ')


data1 = pd.read_excel(r'Y:\Users\guojianan\case\WNBA\2018-WNBA-Rating-GJN.xlsx', sheet_name='Rating change')
data2 = data1.iloc[2:50, 3:].dropna(axis=[0, 1], how='all')
data3 = pd.read_excel(r'Y:\Users\guojianan\case\WNBA\2018-WNBA-Rating-GJN.xlsx', sheet_name='Rating')
dlist = list(data3.iloc[0:12, :1].index)
tlist = [x.date() for x in data2.iloc[:1, :].stack().tolist()]
dt4 = data2.reset_index().iloc[1:, 1:]

ha = dt4.iloc[[(3 * x - 1) for x in range(1, 13)], :1]
ha2 = ha.stack().unstack(0)
ha2.columns = dlist

sp = dt4.iloc[[(3 * x - 2) for x in range(1, 13)], 1:].stack().unstack(0)
sp.index = tlist
sp.columns = dlist

hd = dt4.iloc[[(3 * x - 3) for x in range(1, 13)], 1:].stack().unstack(0)
hd.index = tlist
hd.columns = dlist
ddic = {}
for i in dlist:
    ddic[i] = Float()
ha.columns = ['HA']
ha.index = dlist
ha = ha.reset_index()
ha.to_sql('2018ha', engine, schema='wnba', if_exists='replace', index=False, dtype={'HA': Float()})
hd = hd.reset_index()
sp = sp.reset_index()
ddic['index'] = Date()
hd.to_sql('2018handicap', engine, schema='wnba', if_exists='replace', index=False, dtype=ddic)
sp.to_sql('2018hilo', engine, schema='wnba', if_exists='replace', index=False, dtype=ddic)