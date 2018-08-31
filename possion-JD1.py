# coding=utf-8 #
# Author GJN #

import numpy as np
import scipy.stats as st
import matplotlib.pyplot as plt
import xlwings as xw
import sys
from PyQt5 import QtWidgets
import Rfunc

sup = -0.19
ttg = 3.09
dadj = 0.07
dsplit = 0.5
margin = {'had': 112.5, 'ttg': 131, 'crs': 149, 'hilo': 108.5}


def matrix(sup, ttg, dadj=0.07, dsplit=0.5):
    home = {}
    away = {}
    for i in range(13):
        home[i] = st.poisson.pmf(i, (ttg + sup) / 2)
        away[i] = st.poisson.pmf(i, (ttg - sup) / 2)

    H, A, D = 0, 0, 0
    for i in home:
        for j in away:
            x = home[i] * away[j]
            if i > j:
                H += x
            elif i == j:
                D += x
            elif i < j:
                A += x

    dx = 1 + dadj
    hx = (H - D * dadj * dsplit) / H
    ax = (A - D * dadj * (1 - dsplit)) / A

    score = {}
    for i in home:
        for j in away:
            x = home[i] * away[j]
            if i > j:
                score[(i, j)] = x * hx
            elif i == j:
                score[(i, j)] = x * dx
            elif i < j:
                score[(i, j)] = x * ax
    return score


def had(score):
    H, A, D = 0, 0, 0
    for s in score:
        if s[0] > s[1]:
            H += score[s]
        elif s[0] == s[1]:
            D += score[s]
        elif s[0] < s[1]:
            A += score[s]
    return {'H': H, 'A': A, 'D': D}


def hilo(score, line=2.5):
    Hi, Lo = 0, 0
    for s in score:
        if s[0] + s[1] > line:
            Hi += score[s]
        elif s[0] + s[1] < line:
            Lo += score[s]
    return {'Hi': Hi, 'Lo': Lo}


def margin_ui(result):
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Rfunc.Ui_MainWindow()
    ui.setupUi(MainWindow)
    ui.main()
    ui.had_init((result[0], result[1], result[2]))
    ui.hilo_init((result[6], result[7]))
    MainWindow.show()
    app.exec_()
    return ui.op


wb = xw.Book('Price-2018.xlsx')
sh = wb.sheets[0]
sh.range((13, 1), (13, 12)).value = ['Home', 'Away', 'H',  'D', 'A', 'H',  'D', 'A', 'Hi', 'Lo', 'Hi', 'Lo']
for i in range(2, 11):
    if sh.range(i, 1).value is not None:
        score = matrix(sh.range(i, 11).value, sh.range(i, 12).value)
        HAD = had(score)
        HILO = hilo(score)
        home = sh.range(i, 1).value
        away = sh.range(i, 5).value
        j = i + 12
        H, A, D, Hi, Lo = HAD['H'],  HAD['A'],  HAD['D'], HILO['Hi'], HILO['Lo']
        # mhad = (margin['had'] / 100) / ((H * H) + (D * D) + (A * A))
        # mhilo = (margin['hilo'] / 100) / ((Hi * Hi) + (Lo * Lo))
        # result = [1 / (mhad * H * H), 1 / (mhad * D * D), 1 / (mhad * A * A), 1 / (mhilo * Hi * Hi), 1 / (mhilo * Lo * Lo),
        #     1 / HAD['H'], 1 / HAD['D'], 1 / HAD['A'], 1 / HILO['Hi'], 1 / HILO['Lo']]
        mhad = 100 / margin['had']
        mhilo = 100 / margin['hilo']
        result = [1 / HAD['H'], 1 / HAD['D'], 1 / HAD['A'], mhad / HAD['H'], mhad / HAD['D'], mhad / HAD['A'],
                  1 / HILO['Hi'], 1 / HILO['Lo'], mhilo / HILO['Hi'], mhilo / HILO['Lo']]
        r2 = [x.round(3) for x in result]
        sh.range((j, 1), (j, 2)).value = [home, away]
        sh.range((j, 3), (j, 12)).value = r2
        r3 = margin_ui(result)
        sh.range((i, 13), (i, 15)).value = r3[:3]
        sh.range((i, 17), (i, 18)).value = r3[3:]



