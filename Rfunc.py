# coding=utf-8 #
# Author GJN #

import numpy as np
import scipy.stats as st
import matplotlib.pyplot as plt
import xlwings as xw
from scipy.optimize import fsolve
from PyQt5 import QtCore, QtGui, QtWidgets

sup = -0.19
ttg = 3.09
dadj = 0.07
dsplit = 0.5
margin = {'had': 112.49, 'ttg': 131, 'crs': 149, 'hilo': 108.49}


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
    return [H, A, D]
    # return {'H': H, 'A': A, 'D': D}


def hilo(score, line=2.5):
    Hi, Lo = 0, 0
    for s in score:
        if s[0] + s[1] > line:
            Hi += score[s]
        elif s[0] + s[1] < line:
            Lo += score[s]
    return {'Hi': Hi, 'Lo': Lo}


def line(score, line):

    def over(x):
        o = 0
        for s in score:
            if s[0] - s[1] > -x:
                o += score[s]
        return o

    def under(x):
        o = 0
        for s in score:
            if s[0] - s[1] < -x:
                o += score[s]
        return o

    if line % 1 == 0.5:
        return [over(line), under(line)]
    if line % 1 == 0:
        x = over(line)
        y = under(line)
        z = x + y
        return [x / z, y / z]
    if line % 1 == 0.25:
        x = over(line - 0.25)
        y = under(line - 0.25)
        z = x + y
        u = y / z
        v = under(line + 0.25)
        w = 2 / (1 / u + 1 / v)
        return [1 - w, w]
    if line % 1 == 0.75:
        x = over(line + 0.25)
        y = under(line + 0.25)
        z = x + y
        u = x / z
        v = over(line - 0.25)
        w = 2 / (1 / u + 1 / v)
        return [w, 1 - w]


def hline(score, line):
    def over(x):
        o = 0
        for s in score:
            if s[0] + s[1] > x:
                o += score[s]
        return o

    def under(x):
        o = 0
        for s in score:
            if s[0] + s[1] < x:
                o += score[s]
        return o

    if line % 1 == 0.5:
        return [over(line), under(line)]
    if line % 1 == 0:
        x = over(line)
        y = under(line)
        z = x + y
        return [x / z, y / z]
    if line % 1 == 0.25:
        x = over(line - 0.25)
        y = under(line - 0.25)
        z = x + y
        u = y / z
        v = under(line + 0.25)
        w = 2 / (1 / u + 1 / v)
        return [1 - w, w]
    if line % 1 == 0.75:
        x = over(line + 0.25)
        y = under(line + 0.25)
        z = x + y
        u = x / z
        v = over(line - 0.25)
        w = 2 / (1 / u + 1 / v)
        return [w, 1 - w]


def func(i, l1, l2, n1, n2):
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

    score = matrix(i[0], i[1])

    def line(score, line):

        def over(x):
            o = 0
            for s in score:
                if s[0] - s[1] > -x:
                    o += score[s]
            return o

        def under(x):
            o = 0
            for s in score:
                if s[0] - s[1] < -x:
                    o += score[s]
            return o

        if line % 1 == 0.5:
            return [over(line), under(line)]
        if line % 1 == 0:
            x = over(line)
            y = under(line)
            z = x + y
            return [x / z, y / z]
        if line % 1 == 0.25:
            x = over(line - 0.25)
            y = under(line - 0.25)
            z = x + y
            u = y / z
            v = under(line + 0.25)
            w = 2 / (1 / u + 1 / v)
            return [1 - w, w]
        if line % 1 == 0.75:
            x = over(line + 0.25)
            y = under(line + 0.25)
            z = x + y
            u = x / z
            v = over(line - 0.25)
            w = 2 / (1 / u + 1 / v)
            return [w, 1 - w]

    def hline(score, line):
        def over(x):
            o = 0
            for s in score:
                if s[0] + s[1] > x:
                    o += score[s]
            return o

        def under(x):
            o = 0
            for s in score:
                if s[0] + s[1] < x:
                    o += score[s]
            return o

        if line % 1 == 0.5:
            return [over(line), under(line)]
        if line % 1 == 0:
            x = over(line)
            y = under(line)
            z = x + y
            return [x / z, y / z]
        if line % 1 == 0.25:
            x = over(line - 0.25)
            y = under(line - 0.25)
            z = x + y
            u = y / z
            v = under(line + 0.25)
            w = 2 / (1 / u + 1 / v)
            return [1 - w, w]
        if line % 1 == 0.75:
            x = over(line + 0.25)
            y = under(line + 0.25)
            z = x + y
            u = x / z
            v = over(line - 0.25)
            w = 2 / (1 / u + 1 / v)
            return [w, 1 - w]

    nh = np.array(line(score, l1))
    nhl = np.array(hline(score, l2))
    nxl = np.array([nh[0], nhl[0]])
    nx = np.array([n1, n2])

    return (nxl - nx).tolist()


def rarg(x):
    return x[0], x[2], 1 / x[1], 1 / x[3]


# 输入(ha-line, ha-home, ha-line, hilo-line, hilo-home, hilo-away),输出（sup, ttg)
def tost(x):
    hilo = x[4] * (1 / x[4] + 1 / x[5])
    try:
        hline = -float(x[0].split('/')[0])
    except AttributeError:
        hline = -x[0]
    handi = x[1] * (1 / x[1] + 1 / x[2])
    arg = (hline, handi, x[3], hilo)
    x0 = 0
    delta = 0
    na = np.array([x0, 2])
    f = fsolve(func, na, rarg(arg))
    # 180923部分大sup和ttg解不出来，优化了初始值
    while (f == na).all():
        delta += 1
        x1 = x0 + delta
        na = np.array([x1, 2])
        f = fsolve(func, np.array(na), rarg(arg))
        if (f == na).all():
            x1 = x0 - delta
            na = np.array([x1, 2])
            f = fsolve(func, np.array(na), rarg(arg))
    # if (f == na).all():
    #     x0 = 1
    #     na = np.array([x0, 3.5])
    #     f = fsolve(Rfunc.func, na, Rfunc.rarg(arg))
    #     while (f == na).all():
    #         x0 += 1
    #         na = np.array([x0, 3.5])
    #         f = fsolve(Rfunc.func, np.array([0, 3.5]), Rfunc.rarg(arg))
    return f


x = (0.5, 1.86, 2.5, 2.05)
f = fsolve(func, np.array([0, 2]), rarg(x))


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(632, 429)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayoutWidget = QtWidgets.QWidget(self.centralwidget)
        self.gridLayoutWidget.setGeometry(QtCore.QRect(30, 70, 295, 121))
        self.gridLayoutWidget.setObjectName("gridLayoutWidget")
        self.gridLayout = QtWidgets.QGridLayout(self.gridLayoutWidget)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)
        self.gridLayout.setObjectName("gridLayout")
        self.label_3 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_3.setObjectName("label_3")
        self.gridLayout.addWidget(self.label_3, 0, 2, 1, 1)
        self.lineEdit = QtWidgets.QLineEdit(self.gridLayoutWidget)
        self.lineEdit.setObjectName("lineEdit")
        self.gridLayout.addWidget(self.lineEdit, 2, 0, 1, 1)
        self.lineEdit_2 = QtWidgets.QLineEdit(self.gridLayoutWidget)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.gridLayout.addWidget(self.lineEdit_2, 2, 1, 1, 1)
        self.pushButton_2 = QtWidgets.QPushButton(self.gridLayoutWidget)
        self.pushButton_2.setObjectName("pushButton_2")
        self.gridLayout.addWidget(self.pushButton_2, 3, 1, 1, 1)
        self.pushButton = QtWidgets.QPushButton(self.gridLayoutWidget)
        self.pushButton.setObjectName("pushButton")
        self.gridLayout.addWidget(self.pushButton, 3, 0, 1, 1)
        self.pushButton_3 = QtWidgets.QPushButton(self.gridLayoutWidget)
        self.pushButton_3.setObjectName("pushButton_3")
        self.gridLayout.addWidget(self.pushButton_3, 3, 2, 1, 1)
        self.lineEdit_3 = QtWidgets.QLineEdit(self.gridLayoutWidget)
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.gridLayout.addWidget(self.lineEdit_3, 2, 2, 1, 1)
        self.label = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 0, 0, 1, 1)
        self.label_2 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_2.setObjectName("label_2")
        self.gridLayout.addWidget(self.label_2, 0, 1, 1, 1)
        self.label_6 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_6.setObjectName("label_6")
        self.gridLayout.addWidget(self.label_6, 1, 0, 1, 1)
        self.label_7 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_7.setObjectName("label_7")
        self.gridLayout.addWidget(self.label_7, 1, 1, 1, 1)
        self.label_8 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_8.setObjectName("label_8")
        self.gridLayout.addWidget(self.label_8, 1, 2, 1, 1)
        self.gridLayoutWidget_2 = QtWidgets.QWidget(self.centralwidget)
        self.gridLayoutWidget_2.setGeometry(QtCore.QRect(380, 70, 221, 121))
        self.gridLayoutWidget_2.setObjectName("gridLayoutWidget_2")
        self.gridLayout_2 = QtWidgets.QGridLayout(self.gridLayoutWidget_2)
        self.gridLayout_2.setContentsMargins(0, 0, 0, 0)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.lineEdit_4 = QtWidgets.QLineEdit(self.gridLayoutWidget_2)
        self.lineEdit_4.setObjectName("lineEdit_4")
        self.gridLayout_2.addWidget(self.lineEdit_4, 2, 0, 1, 1)
        self.pushButton_5 = QtWidgets.QPushButton(self.gridLayoutWidget_2)
        self.pushButton_5.setObjectName("pushButton_5")
        self.gridLayout_2.addWidget(self.pushButton_5, 3, 1, 1, 1)
        self.lineEdit_5 = QtWidgets.QLineEdit(self.gridLayoutWidget_2)
        self.lineEdit_5.setObjectName("lineEdit_5")
        self.gridLayout_2.addWidget(self.lineEdit_5, 2, 1, 1, 1)
        self.label_4 = QtWidgets.QLabel(self.gridLayoutWidget_2)
        self.label_4.setObjectName("label_4")
        self.gridLayout_2.addWidget(self.label_4, 0, 0, 1, 1)
        self.label_5 = QtWidgets.QLabel(self.gridLayoutWidget_2)
        self.label_5.setObjectName("label_5")
        self.gridLayout_2.addWidget(self.label_5, 0, 1, 1, 1)
        self.pushButton_4 = QtWidgets.QPushButton(self.gridLayoutWidget_2)
        self.pushButton_4.setObjectName("pushButton_4")
        self.gridLayout_2.addWidget(self.pushButton_4, 3, 0, 1, 1)
        self.label_9 = QtWidgets.QLabel(self.gridLayoutWidget_2)
        self.label_9.setObjectName("label_9")
        self.gridLayout_2.addWidget(self.label_9, 1, 0, 1, 1)
        self.label_10 = QtWidgets.QLabel(self.gridLayoutWidget_2)
        self.label_10.setObjectName("label_10")
        self.gridLayout_2.addWidget(self.label_10, 1, 1, 1, 1)
        self.pushButton_6 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_6.setGeometry(QtCore.QRect(100, 10, 93, 28))
        self.pushButton_6.setObjectName("pushButton_6")
        self.pushButton_7 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_7.setGeometry(QtCore.QRect(450, 10, 93, 28))
        self.pushButton_7.setObjectName("pushButton_7")
        self.pushButton_8 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_8.setGeometry(QtCore.QRect(360, 240, 231, 101))
        self.pushButton_8.setObjectName("pushButton_8")
        self.textEdit = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit.setGeometry(QtCore.QRect(70, 250, 104, 87))
        self.textEdit.setObjectName("textEdit")
        self.pushButton_9 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_9.setGeometry(QtCore.QRect(200, 280, 93, 28))
        self.pushButton_9.setObjectName("pushButton_9")
        self.label_11 = QtWidgets.QLabel(self.centralwidget)
        self.label_11.setGeometry(QtCore.QRect(140, 210, 72, 15))
        self.label_11.setObjectName("label_11")
        self.label_12 = QtWidgets.QLabel(self.centralwidget)
        self.label_12.setGeometry(QtCore.QRect(450, 210, 72, 15))
        self.label_12.setObjectName("label_12")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 632, 26))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.label_3.setText(_translate("MainWindow", "TextLabel"))
        self.pushButton_2.setText(_translate("MainWindow", "D"))
        self.pushButton.setText(_translate("MainWindow", "H"))
        self.pushButton_3.setText(_translate("MainWindow", "A"))
        self.label.setText(_translate("MainWindow", "TextLabel"))
        self.label_2.setText(_translate("MainWindow", "TextLabel"))
        self.label_6.setText(_translate("MainWindow", "TextLabel"))
        self.label_7.setText(_translate("MainWindow", "TextLabel"))
        self.label_8.setText(_translate("MainWindow", "TextLabel"))
        self.pushButton_5.setText(_translate("MainWindow", "Lo"))
        self.label_4.setText(_translate("MainWindow", "TextLabel"))
        self.label_5.setText(_translate("MainWindow", "TextLabel"))
        self.pushButton_4.setText(_translate("MainWindow", "Hi"))
        self.label_9.setText(_translate("MainWindow", "TextLabel"))
        self.label_10.setText(_translate("MainWindow", "TextLabel"))
        self.pushButton_6.setText(_translate("MainWindow", "PushButton"))
        self.pushButton_7.setText(_translate("MainWindow", "PushButton"))
        self.pushButton_8.setText(_translate("MainWindow", "PushButton"))
        self.pushButton_9.setText(_translate("MainWindow", "PushButton"))
        self.label_11.setText(_translate("MainWindow", "TextLabel"))
        self.label_12.setText(_translate("MainWindow", "TextLabel"))

    def __init__(self):
        self.hi, self.hic, self.lo, self.loc = 0, 0, 0, 0
        self.mhilo = 100 / 108.5
        self.mhad = 100 / 112.5
        self.hlmargin = 108.49
        self.hadmargin = 112.49
        self.h, self.d, self.a = 0, 0, 0
        self.hc, self.dc, self.ac = 0, 0, 0
        self.op = [0, 0, 0, 0, 0]

    def main(self):
        self.pushButton_8.clicked.connect(self.output)

    def output(self):
        self.op = [self.hc, self.dc, self.ac, self.hic, self.loc]
        QtCore.QCoreApplication.quit()

    def had_init(self, had=(2.4, 3.861, 3.083)):
        self.h = round(had[0], 3)
        self.d = round(had[1], 3)
        self.a = round(had[2], 3)
        self.label.setText(str(self.h))
        self.label_2.setText(str(self.d))
        self.label_3.setText(str(self.a))

        self.hc = round(self.h * self.mhad, 3)
        self.dc = round(self.d * self.mhad, 3)
        self.ac = round(self.a * self.mhad, 3)
        self.label_6.setText(str(self.hc))
        self.label_7.setText(str(self.dc))
        self.label_8.setText(str(self.ac))
        self.lineEdit.setText(str(self.hc))
        self.lineEdit_2.setText(str(self.dc))
        self.lineEdit_3.setText(str(self.ac))

        self.hadmargin = round((1 / self.hc + 1 / self.dc + 1 / self.ac) * 100, 3)
        self.label_11.setText(str(self.hadmargin))

        self.lineEdit.editingFinished.connect(self.hchange)
        self.lineEdit_2.editingFinished.connect(self.dchange)
        self.lineEdit_3.editingFinished.connect(self.achange)
        self.lineEdit.textChanged.connect(self.had_margin)
        self.lineEdit_2.textChanged.connect(self.had_margin)
        self.lineEdit_3.textChanged.connect(self.had_margin)

        self.pushButton.clicked.connect(self.hadapt)
        self.pushButton_2.clicked.connect(self.dadapt)
        self.pushButton_3.clicked.connect(self.aadapt)

    def hchange(self):
        try:
            self.hc = float(self.lineEdit.text())
        except:
            self.hc = self.lineEdit.text()
        print('H: ', self.hc)

    def dchange(self):
        try:
            self.dc = float(self.lineEdit_2.text())
        except:
            self.dc = self.lineEdit_2.text()
        print('D: ', self.dc)

    def achange(self):
        try:
            self.ac = float(self.lineEdit_3.text())
        except:
            self.ac = self.lineEdit_3.text()
        print('A: ', self.ac)

    def hadapt(self):
        self.hc = round(int(100 / (1 / self.mhad - 1 / self.dc - 1 / self.ac)) / 100 + 0.01, 2)
        self.lineEdit.setText(str(self.hc))

    def dadapt(self):
        self.dc = round(int(100 / (1 / self.mhad - 1 / self.hc - 1 / self.ac)) / 100 + 0.01, 2)
        self.lineEdit_2.setText(str(self.dc))

    def aadapt(self):
        self.ac = round(int(100 / (1 / self.mhad - 1 / self.hc - 1 / self.dc)) / 100 + 0.01, 2)
        self.lineEdit_3.setText(str(self.ac))

    def had_margin(self):
        try:
            self.hc = float(self.lineEdit.text())
        except:
            self.hc = self.lineEdit.text()
        try:
            self.dc = float(self.lineEdit_2.text())
        except:
            self.dc = self.lineEdit_2.text()
        try:
            self.ac = float(self.lineEdit_3.text())
        except:
            self.ac = self.lineEdit_3.text()
        try:
            self.hadmargin = round((1 / self.hc + 1 / self.dc + 1 / self.ac) * 100, 3)
            self.label_11.setText(str(self.hadmargin))
        except:
            self.label_11.setText('error')

    def hilo_init(self, hl=(2.32, 1.758)):
        self.hi = round(hl[0], 3)
        self.lo = round(hl[1], 3)
        self.label_4.setText(str(self.hi))
        self.label_5.setText(str(self.lo))

        self.hic = round(self.hi * self.mhilo, 3)
        self.loc = round(self.lo * self.mhilo, 3)
        self.label_9.setText(str(self.hic))
        self.label_10.setText(str(self.loc))
        self.lineEdit_4.setText(str(self.hic))
        self.lineEdit_5.setText(str(self.loc))

        self.hlmargin = round((1 / self.hic + 1/ self.loc) * 100, 3)
        self.label_12.setText(str(self.hlmargin))

        self.lineEdit_4.editingFinished.connect(self.hichange)
        self.lineEdit_5.editingFinished.connect(self.lochange)
        self.lineEdit_4.textChanged.connect(self.hilo_margin)
        self.lineEdit_5.textChanged.connect(self.hilo_margin)

        self.pushButton_4.clicked.connect(self.hiadapt)
        self.pushButton_5.clicked.connect(self.loadapt)

    def hichange(self):
        try:
            self.hic = float(self.lineEdit_4.text())
        except:
            self.hic = self.lineEdit_4.text()
        print('high: ', self.hic)

    def lochange(self):
        try:
            self.loc = float(self.lineEdit_5.text())
        except:
            self.loc = self.lineEdit_5.text()
        print('low: ', self.loc)

    def hiadapt(self):
        self.hic = round(int(100 / (1 / self.mhilo - 1 / self.loc)) / 100 + 0.01, 2)
        self.lineEdit_4.setText(str(self.hic))

    def loadapt(self):
        self.loc = round(int(100 / (1 / self.mhilo - 1 / self.hic)) / 100 + 0.01, 2)
        self.lineEdit_5.setText(str(self.loc))

    def hilo_margin(self):
        try:
            self.hic = float(self.lineEdit_4.text())
        except:
            self.hic = self.lineEdit_4.text()
        try:
            self.loc = float(self.lineEdit_5.text())
        except:
            self.loc = self.lineEdit_5.text()
        try:
            self.hlmargin = round((1 / self.hic + 1 / self.loc) * 100, 3)
            self.label_12.setText(str(self.hlmargin))
        except:
            self.label_12.setText('error')


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    ui.main()
    ui.had_init((3.3, 3.3, 3.3))
    ui.hilo_init((2, 2))
    MainWindow.show()
    app.exec_()
    print(ui.hc)
    sys.exit()



