# coding=utf-8 #
# Author GJN #
from datetime import datetime, timedelta
import requests
from bs4 import BeautifulSoup
import pandas as pd
import numpy as np
import xlwings as xw
import os
from scipy.optimize import fsolve
import scipy.stats as st


class BR(object):
    def __init__(self):
        self.league = ''
        self.s = requests.session()
        try:
            # 获取ot
            self.headers = {'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
                                     " (KHTML, like Gecko) Chrome/66.0.3359.181 Safari/537.36"}
            r = self.s.get('https://in.betradar.com/', headers=self.headers)
            soup = BeautifulSoup(r.text, 'lxml')
            ot = soup.find('input', type='hidden')['value']

            # 利用ot和用户名、密码登录
            self.headers = {'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
                                     " (KHTML, like Gecko) Chrome/66.0.3359.181 Safari/537.36",
                       'Origin': 'https://sso.betradar.com',
                       'Referer': 'https://sso.betradar.com/authorize.php?oauth_token=%s' % ot}
            ldata = {'username': 'csl-015', 'password': '2018cslTD', 'login': 'Sign in', 'oauth_token': ot}
            r1 = self.s.post('https://sso.betradar.com/authorize.php', data=ldata, headers=self.headers)
            s1 = BeautifulSoup(r1.text, 'lxml')
            url1 = s1.find('script', type="text/javascript").string.split('\'')[1]
            r2 = self.s.get(url1, headers=self.headers)
        except:
            # 获取ot
            self.headers = {'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
                                     " (KHTML, like Gecko) Chrome/66.0.3359.181 Safari/537.36"}
            r = self.s.get('https://in.betradar.com/', headers=self.headers)
            soup = BeautifulSoup(r.text, 'lxml')
            ot = soup.find('input', type='hidden')['value']

            # 利用ot和用户名、密码登录
            self.headers = {'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
                                     " (KHTML, like Gecko) Chrome/66.0.3359.181 Safari/537.36",
                       'Origin': 'https://sso.betradar.com',
                       'Referer': 'https://sso.betradar.com/authorize.php?oauth_token=%s' % ot}
            ldata = {'username': 'csl-013', 'password': '2018cslTD', 'login': 'Sign in', 'oauth_token': ot}
            r1 = self.s.post('https://sso.betradar.com/authorize.php', data=ldata, headers=self.headers)
            s1 = BeautifulSoup(r1.text, 'lxml')
            url1 = s1.find('script', type="text/javascript").string.split('\'')[1]
            r2 = self.s.get(url1, headers=self.headers)

    def request(self):
        days = int(input('days:'))
        ldic = {'EPL': '1', 'PFL': '52', 'JD1': '82'}
        lname = input('League:')
        self.league = ldic[lname]
        ot = {'Total': 4, 'AHC': 9}
        # pdic = {'EPL': r'E:\Company\League\EPL\\', 'PFL': r'E:\Company\League\PFL\\', 'JD1': r'E:\Company\League\JD1\\'}
        path = os.path.abspath(os.curdir) + '\\archive\\'
        dfs = []
        for o in ot:
            rdata = {'selectedTournaments[]': [self.league], 'selectedBookmakers[]': [459],
                     'startDate': (datetime.now() - timedelta(days=days)).strftime('%d.%m.%Y'), 'endDate': datetime.now().strftime('%d.%m.%Y'),
                     'timeStart': '00:00', 'timeEnd': '23:59',
                     'selOddstype': ot[o]}
            rx = self.s.post('https://in.betradar.com/betradar/SystemMain.php?page=startCloseExportGenerate&noheader=true',
                        data=rdata, headers=self.headers)
            filename = '%s_%s_%s-%s_%s' % (o, 'EPL',
                                           (datetime.now() - timedelta(days=days)).strftime('%m.%d'),
                                           datetime.now().strftime('%m.%d.%H%M'),
                                           'SBO')
            with open(path + "%s.xls" % filename, "wb") as code:
                code.write(rx.content)
            df = pd.read_excel(path + "%s.xls" % filename)
            df2 = df.iloc[1:, [1, 3, 5, 6, 7, 8, 9, 10]]
            o2 = o[0]
            df2.columns = ['H', 'A', o2 + 'SL', o2 + 'CL', o2 + 'SP1', o2 + 'CP1', o2 + 'SP2', o2 + 'CP2']
            dfs.append(df2)

        self.dfx = pd.merge(dfs[0], dfs[1])

        def tost(x):
            hilo = x[1] * (1/x[1] + 1/x[2])
            hline = float(x[3].split('/')[0])
            handi = x[4] * (1/x[4] + 1/x[5])
            arg = (hline, handi, x[0], hilo)
            x0 = 0
            delta = 0
            na = np.array([x0, 2])
            f = fsolve(self.func, na, self.rarg(arg))
            # 180923部分大sup和ttg解不出来，优化了初始值
            while (f == na).all():
                delta += 1
                x1 = x0 + delta
                na = np.array([x1, 2])
                f = fsolve(self.func, np.array(na), self.rarg(arg))
                if (f == na).all():
                    x1 = x0 - delta
                    na = np.array([x1, 2])
                    f = fsolve(self.func, np.array(na), self.rarg(arg))
            # if (f == na).all():
            #     x0 = 1
            #     na = np.array([x0, 3.5])
            #     f = fsolve(Rfunc.func, na, Rfunc.rarg(arg))
            #     while (f == na).all():
            #         x0 += 1
            #         na = np.array([x0, 3.5])
            #         f = fsolve(Rfunc.func, np.array([0, 3.5]), Rfunc.rarg(arg))
            return f

        self.dfx['start'] = self.dfx[['TSL', 'TSP1', 'TSP2', 'ASL', 'ASP1', 'ASP2']].apply(tost, axis=1)
        self.dfx['close'] = self.dfx[['TCL', 'TCP1', 'TCP2', 'ACL', 'ACP1', 'ACP2']].apply(tost, axis=1)

        self.dfx['sta-suo'] = self.dfx['start'].map(lambda x: x[0])
        self.dfx['sta-ttg'] = self.dfx['start'].map(lambda x: x[1])
        self.dfx['clo-suo'] = self.dfx['close'].map(lambda x: x[0])
        self.dfx['clo-ttg'] = self.dfx['close'].map(lambda x: x[1])
        self.dfx.to_csv(path + '%s_%s.csv' % (lname, datetime.now().strftime('%m.%d.%H%M')))

    def linetops(self):
        # df = pd.read_excel(path + 'team name.xlsx')
        # psd = df[['Betradar', 'Soccerway']].set_index('Betradar').to_dict()['Soccerway']
        # psd2 = df[['Betradar', 'PS']].set_index('Betradar').to_dict()['PS']

        wbdic = {'1': 'EPL_Playersheet_2018-2019.xlsm', '52': 'PFL_Player_Sheet_18-19.xlsm'}
        wb = xw.Book(wbdic[self.league])
        bsd = {'HUDDERSFIELD TOWN': 'Huddersfield', 'CRYSTAL PALACE': 'Crystal Palace', 'WATFORD FC': 'Watford', 'EVERTON FC': 'Everton', 'LEICESTER CITY': 'Leicester', 'FULHAM FC': 'Fulham', 'MANCHESTER CITY': 'Manchester City', 'CARDIFF CITY': 'Cardiff', 'BRIGHTON & HOVE ALBION': 'Brighton', 'TOTTENHAM HOTSPUR': 'Tottenham', 'SOUTHAMPTON FC': 'Southampton', 'BURNLEY FC': 'Burnley', 'WOLVERHAMPTON WANDERERS': 'Wolves', 'WEST HAM UNITED': 'West Ham', 'LIVERPOOL FC': 'Liverpool', 'MANCHESTER UNITED': 'Manchester Utd', 'NEWCASTLE UNITED': 'Newcastle', 'CHELSEA FC': 'Chelsea', 'ARSENAL FC': 'Arsenal', 'AFC BOURNEMOUTH': 'Bournemouth',
               'MOREIRENSE FC': 'Moreirense', 'GD CHAVES': 'Chaves', 'CD TONDELA': 'Tondela', 'VITORIA GUIMARAES': 'Vitória Guimarães', 'BOAVISTA FC': 'Boavista', 'PORTIMONENSE SC': 'Portimonense', 'SPORTING BRAGA': 'Sporting Braga', 'FC PORTO': 'Porto', 'CS MARITIMO MADEIRA': 'Marítimo', 'SL BENFICA': 'Benfica', 'CD AVES': 'Desportivo Aves', 'CD FEIRENSE': 'Feirense', 'VITORIA SETUBAL': 'Vitória Setúbal', 'SPORTING CP': 'Sporting CP', 'RIO AVE FC': 'Rio Ave', 'CF BELENENSES': 'Belenenses', 'CD NACIONAL MADEIRA': 'Nacional', 'CD SANTA CLARA': 'Santa Clara'
               }

        bpd = {'CRYSTAL PALACE': 'Crystal_Palace', 'CARDIFF CITY': 'Cardiff', 'MANCHESTER UNITED': 'Man_Utd', 'MANCHESTER CITY': 'Man_City', 'WEST HAM UNITED': 'West_Ham', 'NEWCASTLE UNITED': 'Newcastle', 'ARSENAL FC': 'Arsenal', 'WOLVERHAMPTON WANDERERS': 'Wolves', 'WATFORD FC': 'Watford', 'LEICESTER CITY': 'Leicester', 'FULHAM FC': 'Fulham', 'CHELSEA FC': 'Chelsea', 'EVERTON FC': 'Everton', 'TOTTENHAM HOTSPUR': 'Tottenham', 'BURNLEY FC': 'Burnley', 'SOUTHAMPTON FC': 'Southampton', 'HUDDERSFIELD TOWN': 'Huddersfield', 'BRIGHTON & HOVE ALBION': 'Brighton', 'LIVERPOOL FC': 'Liverpool', 'AFC BOURNEMOUTH': 'Bournemouth'
               , 'MOREIRENSE FC': 'Moreirense', 'GD CHAVES': 'Chaves', 'CD TONDELA': 'Tondela', 'VITORIA GUIMARAES': 'Vitória_Guimarães', 'BOAVISTA FC': 'Boavista', 'PORTIMONENSE SC': 'Portimonense', 'SPORTING BRAGA': 'Braga', 'FC PORTO': 'Porto', 'CS MARITIMO MADEIRA': 'Marítimo', 'SL BENFICA': 'Benfica', 'CD AVES': 'Desportivo_Aves', 'CD FEIRENSE': 'Feirense', 'VITORIA SETUBAL': 'Vitória_Setúbal', 'SPORTING CP': 'Sporting_CP', 'RIO AVE FC': 'Rio_Ave', 'CF BELENENSES': 'Belenenses', 'CD NACIONAL MADEIRA': 'Nacional', 'CD SANTA CLARA': 'Santa_Clara'
               }
        for it in self.dfx.iterrows():
            i = it[1]
            hteam = i[0]
            ateam = i[1]
            at = [hteam, ateam]
            print(at)
            ss = round(i['start'][0], 2)
            cs = round(i['close'][0], 2)
            st = round(i['start'][1], 2)
            ct = round(i['close'][1], 2)
            hl = [float(hhll) for hhll in i['ACL'].split('/')]
            index = [1, -1]
            for ind in range(2):
                ss = index[ind] * ss
                cs = index[ind] * cs
                line = hl[ind]
                sh = wb.sheets[bpd[at[ind]]]
                game = sh.range((10, 1), (80, 4)).value
                r = -1
                for g in game:
                    if g[3] == 'PRL':
                        if g[2] == ['H', 'A'][ind]:
                            if g[1] == bsd[at[1 - ind]]:
                                r = game.index(g) + 10
                sh.range(r, 11).value = line
                sh.range(r, 13).value = cs
                sh.range(r, 15).value = ct
                if abs(ss - cs) >= 0.2:
                    try:
                        sh.range(r, 13).api.AddComment(str(ss))
                    except:
                        sh.range(r, 13).api.Comment.Text(str(ss))
                    sh.range(r, 13).api.Comment.Shape.TextFrame.Characters().Font.Bold = True
                    sh.range(r, 13).api.Comment.Shape.TextFrame.Characters().Font.Name = 'Tahoma'
                if abs(st - ct) >= 0.2:
                    try:
                        sh.range(r, 15).api.AddComment(str(st))
                    except:
                        sh.range(r, 15).api.Comment.Text(str(st))
                    sh.range(r, 15).api.Comment.Shape.TextFrame.Characters().Font.Bold = True
                    sh.range(r, 15).api.Comment.Shape.TextFrame.Characters().Font.Name = 'Tahoma'

    def func(self, i, l1, l2, n1, n2):
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

    def rarg(self, x):
        return x[0], x[2], 1 / x[1], 1 / x[3]


if __name__ == "__main__":
    br = BR()
    br.__init__()
    end = ''
    while end == '':
        br.request()
        print('-------------------------------------------------------------------------------------')
        with pd.option_context('display.max_rows', None, 'display.max_columns', None):
            print(br.dfx.iloc[:, [0, 1, -4, -3, -2, -1]])
        print('-------------------------------------------------------------------------------------')
        tops = input('To PS? (Any Word to quit)')
        print('-------------------------------------------------------------------------------------')
        if tops == '':
            br.linetops()
        end = input('Continue? (Any Word to end)')
        print('-------------------------------------------------------------------------------------')
    end = input('End')






