# coding=utf-8 #
# Author GJN #
import xlrd
import pandas as pd
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import Column, String, Date, Integer, create_engine, ForeignKey, Float
import datetime


# path = r'E:\Company\League\EPL\\'
#
# route = path + 'EPL_Playersheet_2017-2018.xlsm'
# season = '17-18'
# r2 = path + 'EPL_Playersheet_2016-2017.xlsm'
# s2 = '16-17'


def rs(year):
    route = r'E:\Company\League\PFL\PFL_playersheet %s-%s.xlsm' % (year, year + 1)
    season = '%s-%s' % (year, year + 1)
    return route, season


def read_pfl(rs):
    route = rs[0]
    season = rs[1]
    wb = xlrd.open_workbook(route)
    slist = wb.sheet_names()[3:21]

    # df = pd.read_excel(path + 'team name.xlsx')
    # psd = df[['PS', 'Soccerway']].set_index('PS').to_dict()['Soccerway']
    # psd['Brighton_Hove'] = psd['Brighton']
    # psd['Stoke'] = psd['Brighton']
    # plist = [psd[x] for x in slist]

    def read_2017(sn):
        df = pd.read_excel(route, encoding="gb18030", sheet_name=sn, index_col=False)
        df1 = df.iloc[21:, [0, 8, 9, 10, 11, 12, 13, 14, 17, 19,
                                                      20, 21, 22, 23, 24, 25, 26, 27, 28]].dropna(axis=[0, 1], how='all')
        dlist = ['date', 'ateam', 'H/A', 'comp', 'gf', 'ga', 'wdl', 'total', 'Hcap', 'sup', 'vssup', 'ttg', 'vsttg',
                 'totalshotfor', 'totalshootagainst', 'shot ot', 'shot ot a', 'pos', 'o pos']
        df1.columns = dlist
        df1['team'] = sn
        dlist.insert(1, 'team')
        df1['season'] = season
        dlist.append('season')
        df2 = df1[dlist][df1['comp'] == 'PRL']
        h1 = []
        for ha in ['H', 'A']:
            h2 = []
            for chp in ['gf', 'ga','sup', 'ttg']:
                x = df2[df2['H/A'] == ha][chp].mean()
                h2.append(x)
            h2.insert(2, h2[0] + h2[1])
            h2.insert(2, h2[0] - h2[1])
            hadic = {'H': 'Home', 'A': 'Away'}
            h2.insert(0, '%s %s' % (season, hadic[ha]))
            h2.insert(0, sn)
            h1.append(h2)
        data = pd.DataFrame(h1, columns=['team', 'info', 'gf', 'ga', 'gd', 'tg', 'sup', 'ttg'])

        print(sn, slist.index(sn), '/', len(slist), df2.shape)
        return df2, data

    dflist = []
    datalist = []
    for sn in slist:
        x = read_2017(sn)
        dflist.append(x[0])
        datalist.append(x[1])
    df = pd.concat(dflist)
    data = pd.concat(datalist)
    return df, data


def pfltosql(dd):
    df = dd[0]
    data = dd[1]
    df = df.fillna(0)
    df = df.replace(' ', 0)
    df['date'] = df['date'].map(lambda x: datetime.datetime.strptime(x.strftime('%y-%m-%d'), '%d-%m-%y').date())
    Base = declarative_base()
    # engine = create_engine('mysql+mysqlconnector://gjn:pass@172.18.1.158:3306/betradar')
    engine = create_engine('mysql+mysqlconnector://root:password@localhost:3306/pfl')
    dtype = {}
    for c in df.columns:
        i = list(df.columns).index(c)
        if i in [1, 2, 3, 4, 7, 20]:
            dtype[c] = String(25)
        elif i in [9, 10, 11, 12, 13, -2, -1]:
            dtype[c] = Float()
        elif i == 0:
            dtype[c] = Date()
        else:
            dtype[c] = Integer()

    df.to_sql('game', engine, schema='pfl', if_exists='append', index=False, dtype=dtype)

    datype = {}
    for c in df.columns:
        i = list(df.columns).index(c)
        if i in [0, 1]:
            datype[c] = String(25)
        else:
            datype[c] = Float()
    data.to_sql('team', engine, schema='pfl', if_exists='append', index=False, dtype=datype)


# x = read_pfl([r'E:\Company\League\PFL\PFL_Player_Sheet_17-18.xlsm', '17-18'])
# pfltosql(x)
# x = read_pfl(rs(16))
# pfltosql(x)
# x2 = read_pfl(rs(15))
# pfltosql(x2)
x = read_pfl(rs(14))
pfltosql(x)




