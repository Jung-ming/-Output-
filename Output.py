# from 函式放置區 import *
import os
import sys

BASE_DIR = os.path.abspath('C:/Users/m3x06/anaconda3/Lib/site-packages')
sys.path.append(BASE_DIR)

import datetime
import pandas as pd


def 閏年判斷():
    x = datetime.datetime.now()
    x = x.date().year
    if x % 4 == 0 and x % 100 != 0:
        x = '閏年'
        return x
    elif x % 400 == 0:
        x = '閏年'
        return x
    else:
        x = '不是閏年'
        return x


def 日期抓取():
    global 抓取日期
    大月份 = [1, 3, 5, 7, 8, 10, 12]  # 1月31天
    小月份 = [4, 6, 9, 11]  # 1月30天
    # 特殊月份 2月 須注意潤年

    月份 = input('請輸入月份:')
    日期 = input('請輸入日期:')
    總天數 = int(input('請輸入總天數:'))
    抓取日期 = []

    for i in range(總天數):
        合成日期 = 月份 + '/' + 日期
        抓取日期.append(合成日期)
        if int(月份) in 大月份 and int(日期) == 31:
            月份 = str(int(月份) + 1)
            日期 = '1'
        # 第二個不能用if 因為會變成再判斷一次，若不符合又會進入到else，
        # 但第一個 if 若符合後就不該再進入 else 了
        elif int(月份) in 小月份 and int(日期) == 30:
            月份 = str(int(月份) + 1)
            日期 = '1'
        elif int(月份) == 2:
            閏年判斷結果 = 閏年判斷()
            if 閏年判斷結果 == '閏年' and int(日期) == 29:
                月份 = str(int(月份) + 1)
                日期 = '1'
            elif 閏年判斷結果 == '不是閏年' and int(日期) == 28:
                月份 = str(int(月份) + 1)
                日期 = '1'
            else:
                日期 = str(int(日期) + 1)
        else:
            日期 = str(int(日期) + 1)
    print(抓取日期)
    return 抓取日期


def 抓取目標項目(data, 抓取日期):
    目標項目 = set()
    足標記數 = 0
    for i in data['OUTPUT']:
        for 日期 in 抓取日期:
            if 日期 in i:
                目標項目.add(足標記數)
        足標記數 += 1
    足標記數 = 0

    for i in data['DIP首件產出時間/數量']:
        for 日期 in 抓取日期:
            if 日期 in i:
                目標項目.add(足標記數)
        足標記數 += 1
    return 目標項目


def 文件讀取():
    file_name = input('請複製文件名並貼上:')
    # dtype={'開始時間':str} 可以用來設定讀取時的資料型態
    # 避免自動出不想要的結果 (目前不用這麼做了，先刪掉)
    data = pd.read_excel(f'C:/Users/m3x06/PycharmProjects/公司文件處理/公司Output文件/{file_name}'
                         , header=1, sheet_name=['DIP', 'SMT'])
    return data


def 目標選擇():
    print('參考選項\n0. DIP\n1. SMT\n2. DIP+SMT')
    選項 = input('輸入選項')
    print(f'選擇{選項}')
    return 選項


def 處理選定目標與輸出文件(data):
    global 選項
    選項 = int(目標選擇())
    if 選項 == 0:
        data = 文件處理(data['DIP'])
        目標項目 = 抓取目標項目(data, 抓取日期)
        目標項目文件輸出(data, 目標項目)
    elif 選項 == 1:
        data = 文件處理(data['SMT'])
        目標項目 = 抓取目標項目(data, 抓取日期)
        目標項目文件輸出(data, 目標項目)
    elif 選項 == 2:
        # 這裡的想法是選3的話，就是把選項1和2各個跑一遍
        # 所以將選項重置為0，並隨著迴圈變成1和2
        選項 = 0
        所有工作表 = [data['DIP'], data['SMT']]
        for 工作表 in 所有工作表:
            print(f'執行選項{選項}，準備文件處理')
            data = 文件處理(所有工作表[選項])
            print('準備抓取目標項目..')
            目標項目 = 抓取目標項目(data, 抓取日期)
            print('準備輸出文件...')
            目標項目文件輸出(data, 目標項目)
            print('輸出成功')
            選項 += 1


def 文件處理(data):
    # 只留批號1
    data = data[data['批號'] == 1]

    # 丟掉AP和AQ皆為空值的列
    # 先丟AP為空
    data = data[pd.isnull(data['DIP首件產出時間/數量']) == False]
    # 再丟AQ為空
    data = data[pd.isnull(data['OUTPUT']) == False]

    # 去除重複值
    data = data.drop_duplicates(subset=['母工單單號', '名稱規格'])
    # 重置索引，不重置的話，會不方便依照索引去抓取想要的項目
    data = data.reset_index(drop=True)
    return data


def 目標項目文件輸出(data, 目標項目):
    if 選項 == 0:
        輸出工作表 = '輸出結果-DIP.xls'
    elif 選項 == 1:
        輸出工作表 = '輸出結果-SMT.xls'
    data = data.query(f'index == {list(目標項目)}')
    data.to_excel(f'{輸出工作表}', index=False)

    # if 選項 == 0:
    #     輸出檔名 = '輸出結果-DIP.xls'
    # elif 選項 == 1:
    #     輸出檔名 = '輸出結果-SMT.xls'
    # data = data.query(f'index == {list(目標項目)}')
    # data.to_excel(f'{輸出檔名}', index=False, sheet_name='DIP')


data = 文件讀取()
日期抓取()
處理選定目標與輸出文件(data)
