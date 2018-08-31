# coding=utf-8 #
# Author GJN #
import xlrd
import pandas as pd
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import Column, String, Date, Integer, create_engine, ForeignKey, Float

wb = xlrd.open_workbook(r'Y:\Departments\CSLRM\Level D\FB Odds Locals\JD1\JD1__2018.xlsm')
slist = wb.sheet_names()[4:22]


def read_2018(sn):
    df = pd.read_excel(r'Y:\Departments\CSLRM\Level D\FB Odds Locals\JD1\JD1__2018.xlsm', encoding="gb18030",
                       sheet_name=sn, index_col=False)

    df1 = df.dropna(axis=1, how='all').loc[20:].dropna(axis=[0, 1], how='all')
    df1 = df1.iloc[:, [0, 1, 2, 3, 4, 5, 6, 7, 8, 12, 13, 14, 15, 18, 19, 20, 21, 22, 23, 24]].dropna(axis=[0, 1],
                                                                                                      how='all')
    dlist = ['date', 'ateam', 'H/A', 'comp', 'gf', 'ga', 'wdl', 'total', 'goaldif', 'sup', 'vssup', 'ttg', 'vsttg',
             'totalshotfor', 'totalshootagainst', 'shot ot', 'shot ot a', 'pos', 'o pos', 'formation']
    df1.columns = dlist
    df1['team'] = sn
    dlist.insert(1, 'team')
    df2 = df1[dlist][df1['comp'] == 'J1L']

    print(sn, slist.index(sn), '/', len(slist), df2.shape)
    return df2

dflist = []
for sn in slist:
    dflist.append(read_2018(sn))
df = pd.concat(dflist)

Base = declarative_base()
# engine = create_engine('mysql+mysqlconnector://gjn:pass@172.18.1.158:3306/betradar')
engine = create_engine('mysql+mysqlconnector://root:password@localhost:3306/jd1')
dtype = {}
for c in df.columns:
    i = list(df.columns).index(c)
    if i in [1, 2, 3, 4, 7, 20]:
        dtype[c] = String(25)
    elif i in [10, 11, 12, 13]:
        dtype[c] = Float()
    elif i == 0:
        dtype[c] = Date()
    else:
        dtype[c] = Integer()

df.to_sql('2018', engine, schema='jd1', if_exists='replace', index=False, dtype=dtype)


def jd1_2017():
    wb = xlrd.open_workbook('JD1__2017.xlsm')
    slist = wb.sheet_names()[4:22]

    def read_2017(sn):
        df = pd.read_excel('JD1__2017.xlsm', encoding="gb18030", sheet_name=sn, index_col=False)

        df1 = df.dropna(axis=1, how='all').iloc[20:, [0, 5, 6, 7, 8, 9, 10, 11, 12, 16, 17,
                                                      18, 19, 22, 23, 24, 25, 26, 27, 28]].dropna(axis=[0, 1],
                                                                                                  how='all')
        dlist = ['date', 'ateam', 'H/A', 'comp', 'gf', 'ga', 'wdl', 'total', 'goaldif', 'sup', 'vssup', 'ttg', 'vsttg',
                 'totalshotfor', 'totalshootagainst', 'shot ot', 'shot ot a', 'pos', 'o pos', 'formation']
        df1.columns = dlist
        df1['team'] = sn
        dlist.insert(1, 'team')
        df2 = df1[dlist][df1['comp'] == 'J1L']

        print(sn, slist.index(sn), '/', len(slist), df2.shape)
        return df2

    dflist = []
    for sn in slist:
        dflist.append(read_2017(sn))
    df = pd.concat(dflist)

    Base = declarative_base()
    # engine = create_engine('mysql+mysqlconnector://gjn:pass@172.18.1.158:3306/betradar')
    engine = create_engine('mysql+mysqlconnector://root:password@localhost:3306/jd1')
    dtype = {}
    for c in df.columns:
        i = list(df.columns).index(c)
        if i in [1, 2, 3, 4, 7, 20]:
            dtype[c] = String(25)
        elif i in [10, 11, 12, 13]:
            dtype[c] = Float()
        elif i == 0:
            dtype[c] = Date()
        else:
            dtype[c] = Integer()

    df.to_sql('2017', engine, schema='jd1', if_exists='replace', index=False, dtype=dtype)


def jd2_2017():
    wb = xlrd.open_workbook('Japan2 17.xlsm')
    s1 = wb.sheet_names()[0:10]
    slist = wb.sheet_names()[0:11] + wb.sheet_names()[14:25]

    def read_2018(sn):
        df = pd.read_excel('Japan2 17.xlsm', encoding="gb18030", sheet_name=sn, index_col=False)

        df1 = df.dropna(axis=1, how='all').loc[20:].dropna(axis=[0, 1], how='all')
        df1 = df1.iloc[:, [0, 3, 4, 5, 6, 7, 8, 9, 13, 14, 15, 16, 19, 20, 21, 22, 23, 24]].dropna(axis=[0, 1],
                                                                                                   how='all')
        dlist = ['date', 'ateam', 'H/A', 'comp', 'gf', 'ga', 'wdl', 'total', 'sup', 'vssup', 'ttg', 'vsttg',
                 'totalshotfor', 'totalshootagainst', 'shot ot', 'shot ot a', 'pos', 'o pos']
        df1.columns = dlist
        df1['team'] = sn
        dlist.insert(1, 'team')
        df2 = df1.reset_index()[dlist]
        df2 = df2[df2['comp'] == 'J2L'].iloc[:42]

        print(sn, slist.index(sn), '/', len(slist) + 1, df2.shape)
        return df2

    dflist = []
    for sn in slist:
        dflist.append(read_2018(sn))
    df = pd.concat(dflist)

    Base = declarative_base()
    # engine = create_engine('mysql+mysqlconnector://gjn:pass@172.18.1.158:3306/betradar')
    engine = create_engine('mysql+mysqlconnector://root:password@localhost:3306/jd1')
    dtype = {}
    for c in df.columns:
        i = list(df.columns).index(c)
        if i in [1, 2, 3, 4, 7]:
            dtype[c] = String(25)
        elif i in [9, 10, 11, 12]:
            dtype[c] = Float()
        elif i == 0:
            dtype[c] = Date()
        else:
            dtype[c] = Integer()

    df.to_sql('2017jd2', engine, schema='jd1', if_exists='replace', index=False, dtype=dtype)


