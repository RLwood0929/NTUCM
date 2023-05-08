# -*- coning: utf-8 -*-
"""
Created on Fri May 5 19:56:02 2023

@author: Qian

程式功能:匯入excel data檔案
將第一欄的數值與第二欄的數值相加
總和答案存入第三欄
第三欄標題為 總和
將結果輸出為excel檔及json檔案
"""

import pandas as pd
import json

def add0(filename):
    df = pd.read_excel(filename)
    x, y = df.shape
    df.insert(2, column = "總和", value = "0")

    for i in range(x):
        add = df.iloc[i, 0]
        summand = df.iloc[i, 1]
        ans = add + summand
        df.iloc[i, 2] = ans

    df.to_excel("結果.xlsx", index = None, header = True)
    df2 = df.to_json(orient = "records", force_ascii= False)

    with open("結果.json", "w") as file:
        json.dump(df2, file, indent = 4, ensure_ascii = False)
    return "done"