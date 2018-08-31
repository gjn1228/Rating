# coding=utf-8 #
# Author GJN #
import xlrd
import pandas as pd
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import Column, String, Date, Integer, create_engine, ForeignKey, Float


# path = r'E:\Company\League\EPL\\'
#
# route = path + 'EPL_Playersheet_2017-2018.xlsm'
# season = '17-18'
# r2 = path + 'EPL_Playersheet_2016-2017.xlsm'
# s2 = '16-17'


def rs(year):
    route = r'E:\Company\League\EPL\EPL_Playersheet_20%s-20%s.xlsm' % (year, year + 1)
    season = '%s-%s' % (year, year + 1)
    return route, season


def read_epl(rs):
    route = rs[0]
    season = rs[1]
    wb = xlrd.open_workbook(route)
    slist = wb.sheet_names()[4:24]

    # df = pd.read_excel(path + 'team name.xlsx')
    # psd = df[['PS', 'Soccerway']].set_index('PS').to_dict()['Soccerway']
    # psd['Brighton_Hove'] = psd['Brighton']
    # psd['Stoke'] = psd['Brighton']
    # plist = [psd[x] for x in slist]

    def read_2017(sn):
        df = pd.read_excel(route, encoding="gb18030", sheet_name=sn, index_col=False)
        df1 = df.iloc[7:, [0, 8, 9, 10, 11, 12, 13, 14, 17, 19,
                                                      20, 21, 22, 23, 24, 25, 26, 27, 28]].dropna(axis=[0, 1],
                                                                                                  how='all')
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


def epltosql(dd):
    df = dd[0]
    data = dd[1]
    Base = declarative_base()
    # engine = create_engine('mysql+mysqlconnector://gjn:pass@172.18.1.158:3306/betradar')
    engine = create_engine('mysql+mysqlconnector://root:password@localhost:3306/epl')
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

    df.to_sql('game', engine, schema='epl', if_exists='append', index=False, dtype=dtype)

    datype = {}
    for c in df.columns:
        i = list(df.columns).index(c)
        if i in [0, 1]:
            datype[c] = String(25)
        else:
            datype[c] = Float()
    data.to_sql('team', engine, schema='epl', if_exists='append', index=False, dtype=datype)


# epltosql(read_epl(rs(17)))
# epltosql(read_epl(rs(16)))
# epltosql(read_epl(rs(15)))
# epltosql(read_epl(rs(14)))
# x = read_epl(rs(13))
# epltosql(x)

def epl2018player():
    route = rs(18)[0]
    wb = xlrd.open_workbook(route)
    slist = wb.sheet_names()[:20]
    glist = [58, 56, 58, 53, 61, 65, 55, 59, 51, 56, 58, 60, 58, 58, 57, 56, 57, 62, 58, 61]
    dlist = []
    for sn in slist:
        gr = glist[slist.index(sn)]
        st = wb.sheet_by_name(sn)

        def getlist(row):
            x = st.row(row)[29:gr]
            y = [xx.value for xx in x]
            y = [int(yy) if type(yy) == float else yy for yy in y]
            y = [0 if yy in ['-', ''] else yy for yy in y]
            return y

        name = getlist(7)
        pos = getlist(6)
        age = getlist(8)
        NO = getlist(5)
        col = list(range(30, gr + 1))
        df = pd.DataFrame({'team': sn, 'name': name, 'position': pos, 'age': age, 'NO': NO, 'col': col})
        dlist.append(df)

    df = pd.concat(dlist)

    engine = create_engine('mysql+mysqlconnector://root:password@localhost:3306/epl')
    dtype = {}
    for c in df.columns:
        i = list(df.columns).index(c)
        if i in [0, 1, 2]:
            dtype[c] = String(25)
        else:
            dtype[c] = Integer()

    df.to_sql('2018player', engine, schema='epl', if_exists='replace', index=False, dtype=dtype)


