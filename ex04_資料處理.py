import datetime


def 抓取目標項目(data, 抓取日期):
    目標項目 = set()
    for index, value in enumerate(data['OUTPUT']):
        for 日期 in 抓取日期:
            if 日期 in value:
                目標項目.add(index)
    for index, value in enumerate(data['DIP首件產出時間/數量']):
        for 日期 in 抓取日期:
            if 日期 in value:
                目標項目.add(index)
    return 目標項目


def 目標項目與資料比對(data, 目標項目):
    if type(data) is list:
        for index, value in enumerate(data):
            # 目標項目是set 不能像list一樣用足標去訪問
            if index == 0:
                data_DIP = value.query(f'index == {list(目標項目[index])}')
            else:
                data_SMT = value.query(f'index == {list(目標項目[index])}')
    else:
        data = data.query(f'index == {list(目標項目)}')
        data.to_excel('輸出結果.xls', index=False)

    return data_DIP, data_SMT


def 抓取Output足標(data, 抓取日期):
    目標項目 = set()
    for 足標, value in enumerate(data['OUTPUT']):
        if 抓取日期 in value:
            目標項目.add(足標)
    return 目標項目


def 抓取DIP首件足標(data, 抓取日期):
    目標項目 = set()
    for 足標, value in enumerate(data['DIP首件產出時間/數量']):
        if 抓取日期 in value:
            目標項目.add(足標)
    return 目標項目


def 排序資料(data):
    # 將'TEST'和'成品'皆為X的部分排序到最前面
    重置足標_雙X = []
    重置足標_一般 = []
    for index, row in data.iterrows():
        if row['TEST'] == 'X' and row['成品'] == 'X':
            重置足標_雙X.append(index)
        else:
            重置足標_一般.append(index)
    重置足標 = 重置足標_雙X + 重置足標_一般
    data = data.reindex(重置足標)

    return data


def 標記類別(data, 足標, 類別值):
    if 足標 == None:
        pass
    else:
        for 足標值 in 足標:
            data.at[足標值, '類別'] = 類別值
    return data


def 日期格式與排序的類別標示(抓取日期, data):
    # 將不同日期標為不同類別，以便後續排序和更改指定格式
    # 足標需區分成AP與AQ，因為兩種欄位格式不同
    黃底日期 = 抓取日期[-1]
    黃底日期足標_AP = 抓取DIP首件足標(data, 黃底日期)
    黃底日期足標_AQ = 抓取Output足標(data, 黃底日期)

    抓取日期.remove(黃底日期)

    當天日期_日期格式 = datetime.date.today()
    當天日期_文字格式 = 當天日期_日期格式.strftime('%#m/%#d')
    # 當天日期_文字格式 = '5/11'
    當天日期足標_AP = 抓取DIP首件足標(data, 當天日期_文字格式)
    當天日期足標_AQ = 抓取Output足標(data, 當天日期_文字格式)
    抓取日期.remove(當天日期_文字格式)

    if len(抓取日期) != 0:
        for i in range(len(抓取日期)):
            # i會從0開始，因此首先建立個別的AP和AQ足標
            # 接下來另外建立待合併的AP和AQ足標，並將其與最初的AP和AQ足標合併
            if i == 0:
                剩餘日期足標_AP = 抓取DIP首件足標(data, 抓取日期[i])
                剩餘日期足標_AQ = 抓取Output足標(data, 抓取日期[i])
            if i != 0:
                剩餘日期足標_待合併AP = 抓取DIP首件足標(data, 抓取日期[i])
                剩餘日期足標_待合併AQ = 抓取Output足標(data, 抓取日期[i])
                剩餘日期足標_AP = 剩餘日期足標_AP | 剩餘日期足標_待合併AP
                剩餘日期足標_AQ = 剩餘日期足標_AQ | 剩餘日期足標_待合併AQ

    else:
        剩餘日期足標_AP = None
        剩餘日期足標_AQ = None

    data = data.assign(類別='')
    # 將全部黃底日期標上類別
    全黃底日期足標 = 黃底日期足標_AP | 黃底日期足標_AQ
    data = 標記類別(data, 全黃底日期足標, 3)

    # 將全部剩餘日期標上類別
    if 剩餘日期足標_AP != None:
        全剩餘日期 = 剩餘日期足標_AP | 剩餘日期足標_AQ
        for 足標 in 全剩餘日期:
            data.at[足標, '類別'] = 2

    # 將全部當天日期標上類別
    全當天日期足標 = 當天日期足標_AP | 當天日期足標_AQ
    data = 標記類別(data, 全當天日期足標, 1)

    data = data.sort_values(by='類別', ascending=True)

    # 前面已按類別順序1~3排序，接下來再按把雙X的項目和其他項目分開
    # 由於是按照順序逐一檢查並分開，所以並不影響到類別的排序
    data = 排序資料(data)

    抓取日期.append(當天日期_文字格式)
    抓取日期.append(黃底日期)
    return 黃底日期足標_AP, 黃底日期足標_AQ, 當天日期足標_AP, 當天日期足標_AQ, 剩餘日期足標_AP, 剩餘日期足標_AQ, data
