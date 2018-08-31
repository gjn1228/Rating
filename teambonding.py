# coding=utf-8 #
# Author GJN #

import pandas as pd
import numpy as np

df = pd.read_excel(r'Y:\Users\guojianan\case\teambonding\2018-07.xlsx', header=0, encoding="gb18030")
df2 = df.reset_index()
df3 = df2.iloc[[0,1,2,4,5,6], [4,5,6]]
df4 = df3.stack().unstack(0)
dic = df4.to_dict()
A, B, C, D = [0, 0, 0, 0]
# x = np.array([0, 0, 0, 0])

allc = []
for i0 in dic[0]:
    for i1 in dic[1]:
        for i2 in dic[2]:
            for i4 in dic[4]:
                for i5 in dic[5]:
                    for i6 in dic[6]:
                        a, b, c, d = [0, 0, 0, 0]
                        if i0 == 'H':
                            a += 3
                        elif i0 == 'A':
                            c += 3
                        else:
                            a += 1
                            c += 1

                        if i1 == 'H':
                            a += 3
                        elif i1 == 'A':
                            d += 3
                        else:
                            a += 1
                            d += 1

                        if i2 == 'H':
                            a += 3
                        elif i2 == 'A':
                            b += 3
                        else:
                            a += 1
                            b += 1

                        if i4 == 'H':
                            b += 3
                        elif i0 == 'A':
                            d += 3
                        else:
                            b += 1
                            d += 1

                        if i5 == 'H':
                            b += 3
                        elif i5 == 'A':
                            c += 3
                        else:
                            b += 1
                            c += 1

                        if i6 == 'H':
                            c += 3
                        elif i6 == 'A':
                            d += 3
                        else:
                            c += 1
                            d += 1
                        g = [a, b, c, d]
                        high = max(g)
                        chp = []
                        for x in g:
                            if x == high:
                                chp.append(x)
                        allc.append(chp)
                        prob = dic[0][i0] * dic[1][i1] * dic[2][i2] * dic[4][i4] * dic[5][i5] * dic[6][i6]
                        l = len(chp)
                        if a in chp:
                            A += prob / l
                        if b in chp:
                            B += prob / l
                        if c in chp:
                            C += prob / l
                        if d in chp:
                            D += prob / l

print(A)
print(B)
print(C)
print(D)











