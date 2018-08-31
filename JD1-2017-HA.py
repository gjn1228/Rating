# coding=utf-8 #
# Author GJN #
import xlrd
import pandas as pd
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import Column, String, Date, Integer, create_engine, ForeignKey, Float

# wb = xlrd.open_workbook(r'Y:\Departments\CSLRM\Level D\FB Odds Locals\JD1\JD1__2018.xlsm')

Base = declarative_base()
# engine = create_engine('mysql+mysqlconnector://gjn:pass@172.18.1.158:3306/betradar')
engine = create_engine('mysql+mysqlconnector://root:password@localhost:3306/jd1')


def get_ha(table):
    df = pd.read_sql(table, engine)
    dfh = df[df['H/A'] == 'H'].iloc[:, [1, 3, 9, 10]].groupby('team').sum()
    dfa = df[df['H/A'] == 'A'].iloc[:, [1, 3, 9, 10]].groupby('team').sum()
    dfj = dfh - dfa
    row = dfj.shape[0] - 1
    s = dfj['sup'].sum() / (2 * row)
    ha = dfj['sup'].map(lambda x: (x - s) / (row - 1))
    return ha


ha2017 = get_ha('2017')
hajd2 = get_ha('2017jd2')


