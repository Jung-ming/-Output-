from ex04_資料處理 import *
from ex03_日期區間 import 取得日期區間
import platform
import os


def 獲取桌面路徑():
    # 獲取使用者電腦系統 windows Linux等
    系統 = platform.system()
    目錄 = os.path.expanduser("~")

    if 系統 == "Windows":
        return os.path.join(目錄, "Desktop").replace('\\', '/')
    elif 系統 == "Darwin":  # macOS
        return os.path.join(目錄, "Desktop").replace('\\', '/')
    elif 系統 == "Linux":
        return os.path.join(目錄, "Desktop").replace('\\', '/')
    else:
        # 默認返回用戶主目錄
        return 目錄.replace('\\', '/')


def 文件讀取(文件路徑):
    # dtype={'開始時間':str} 可以用來設定讀取時的資料型態
    # 避免自動出不想要的結果 (目前不用這麼做了，先刪掉)
    # 利用tinker讓使用者自行選擇想要操作的文件
    data = pd.read_excel(f'{文件路徑}', header=1, sheet_name=['DIP', 'SMT'])
    data_DIP = data['DIP']
    data_SMT = data['SMT']

    return data_DIP, data_SMT


def 資料預處理(data):
    # 丟掉AP和AQ皆為空值的列
    data = data.dropna(subset=['DIP首件產出時間/數量', 'OUTPUT'])
    # 去除重複值
    data = data.drop_duplicates(subset=['母工單單號', '名稱規格', 'OUTPUT'], ignore_index=True)

    return data


def 格式更改(抓取日期, data):
    # 改變日期讀取有2種寫法
    # datetime_format='mm/dd yyyy
    # date_format='mmmm dd yyyy'
    # 使用上取決於資料本身的狀況，如果資料包含時間 則用datetime_format
    # 如果只有年月日則使用date_format

    # 獲取當天日期，並將其轉換成文字
    # 當天日期_文字格式 = 當天日期_日期格式.strftime('%#m/%#d')
    當天日期_日期格式 = datetime.date.today()
    當天日期_文字格式 = 當天日期_日期格式.strftime('%m%d')
    輸出檔名 = 'Output輸出檔' + 當天日期_文字格式 + '.xlsx'
    桌面路徑 = 獲取桌面路徑()
    writer = pd.ExcelWriter(f'{桌面路徑}/{輸出檔名}', engine='xlsxwriter', datetime_format='mm/dd hh:mm')

    # for 工作表 in data:
    #     for index, value in enumerate(工作表['生管備註']):
    #         formatted_date = value.strftime("%#m/%#d")
    #         工作表.at[index, '生管備註'] = formatted_date
    #     for index, value in enumerate(工作表['開始時間']):
    #         formatted_date = value.strftime("%#m/%#d %H:%M:%S")
    #         工作表.at[index, '開始時間'] = formatted_date
    #     for index, value in enumerate(工作表['結束時間']):
    #         formatted_date = value.strftime("%#m/%#d %H:%M:%S")
    #         工作表.at[index, '結束時間'] = formatted_date

    # 一些設定
    列寬_10 = [0, 1, 9]
    列寬_5 = [43, 44, 46]
    隱藏行 = ['D', 'F', 'H', 'I', 'L', 'O', 'Q', 'R', 'S',
           'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB',
           'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ',
           'AM', 'AN', 'AO']

    # 設置底色
    # 假日用-淺藍 #b7dee8 、 當日用-淺綠 #92d050 、 明天用-黃色 #ffff00
    # 此部分注意!!! 原本僅用底色格式時，會造成換行格式消失，也就是單元格內若有多行資料就會全部變成一行
    # 因此加入'text_wrap': True處理
    # 經過ChatGPT說明，會有此現象是因為，套用格式時導致換行字元\n被忽略
    # 加入'text_wrap': True確實是一個解法
    淺綠格式 = writer.book.add_format({'bg_color': '#92d050', 'valign': 'vcenter', 'font_size': 10, 'text_wrap': True})
    淺藍格式 = writer.book.add_format({'bg_color': '#b7dee8', 'valign': 'vcenter', 'font_size': 10, 'text_wrap': True})
    黃色格式 = writer.book.add_format({'bg_color': '#ffff00', 'valign': 'vcenter', 'font_size': 10, 'text_wrap': True})
    日期格式_置中 = writer.book.add_format({'align': 'center', 'valign': 'vcenter', 'font_size': 10})
    日期格式_靠左 = writer.book.add_format({'valign': 'vcenter', 'font_size': 10})
    # 若要設置標題格式，則需要先用下方程式碼消除標題單元格格式
    pandas.io.formats.excel.ExcelFormatter.header_style = None
    # 設置格式
    # 字體 'font-family': 'Times New Roman'
    # 大小 'font-size': '12pt'
    # 'text-align': 'center'
    # 'vertical-align': 'middle'

    # 將DataFrame寫入Excel
    工作表名稱 = ['DIP', 'SMT']

    for index, 工作表 in enumerate(data):
        # 這裡的len不會算到標題行，所以加1
        # 方便後續操作有考慮到標題行的狀況
        總筆數 = len(工作表.index) + 1

        工作表.to_excel(writer, index=False, sheet_name=工作表名稱[index])

        # 獲取工作表
        worksheet = writer.sheets[工作表名稱[index]]
        # 用for迴圈跑過所有Output
        # data.iloc[行數-1, 42] 貌似會直接跳過標題行，從內容的第一格開始，並且足標為0
        # 這邊所抓的足標是從內容的第一行開始，並且足標為0，所以不需要像前面-1，而是一開始的行設定要+1
        # worksheet.write(行足標, 41, 工作表.iloc[內容的行足標(不包含標題行), 41], 黃色格式)
        # 就目前的邏輯來看，最後一個日期依定會是要填入黃色的日期，故使用list的pop()函數
        黃底日期足標_AP, 黃底日期足標_AQ, 當天日期足標_AP, 當天日期足標_AQ, 剩餘日期足標_AP, 剩餘日期足標_AQ, _ \
            = 日期格式與排序的類別標示(抓取日期, 工作表)
        for 足標 in 黃底日期足標_AP:
            worksheet.write(足標 + 1, 41, 工作表.iloc[足標, 41], 黃色格式)
        for 足標 in 黃底日期足標_AQ:
            worksheet.write(足標 + 1, 42, 工作表.iloc[足標, 42], 黃色格式)

        for 足標 in 當天日期足標_AP:
            worksheet.write(足標 + 1, 41, 工作表.iloc[足標, 41], 淺綠格式)
        for 足標 in 當天日期足標_AQ:
            worksheet.write(足標 + 1, 42, 工作表.iloc[足標, 42], 淺綠格式)

        if 剩餘日期足標_AP != None:
            for 足標 in 剩餘日期足標_AP:
                worksheet.write(足標 + 1, 41, 工作表.iloc[足標, 41], 淺藍格式)
            for 足標 in 剩餘日期足標_AQ:
                worksheet.write(足標 + 1, 42, 工作表.iloc[足標, 42], 淺藍格式)

        # .set_column(0, 0, 10) 用來設置列寬，3個參數分別為，起始列、結束列和列寬
        for 列 in 列寬_10:
            worksheet.set_column(列, 列, 12)
        for 列 in 列寬_5:
            worksheet.set_column(列, 列, 5)
        # DIP首件設置16
        worksheet.set_column(41, 41, 16)
        # Output設置30
        worksheet.set_column(42, 42, 30)
        # 生管備註設置9
        worksheet.set_column(45, 45, 9)
        # 批號
        worksheet.set_column(2, 2, 2)
        # 出足數
        worksheet.set_column(15, 15, 2)
        # 工令量、排產量
        worksheet.set_column(10, 10, 4)
        worksheet.set_column(12, 12, 4)

        # 開始時間、結束時間
        worksheet.set_column(36, 36, 10)
        worksheet.set_column(37, 37, 10)

        # 名稱規格
        worksheet.set_column(4, 4, 25)

        # 隱藏設置
        for 隱藏 in 隱藏行:
            worksheet.set_column(f'{隱藏}:{隱藏}', None, None, {'hidden': True})

        標題行格式 = writer.book.add_format({'align': 'center',
                                        'font_size': 10,
                                        'valign': 'vcenter',
                                        'text_wrap': True})

        一般行格式 = writer.book.add_format({
            'font_size': 10,
            'valign': 'vcenter',
            'text_wrap': True})

        for 行數 in range(總筆數):
            if 行數 != 0:
                worksheet.set_row(行數, 40, 一般行格式)
            else:
                worksheet.set_row(行數, 56, 標題行格式)
        # for 行數 in range(總筆數 - 1):
        #     # 就目前經驗來講，操作上最好都讓總筆數=內容行數，不包含標題行，這樣便能直接按照如下邏輯撰寫
        #     # 這裡的標題行-1 變為101，也就是內容的總筆數
        #     # 而若從內容開始，則第一個參數+1，才是寫道內容的第一行
        #     # 工作表.iloc則沿用原數字，因為該函數本就是以內容的足標為主，不包含標題行
        #     worksheet.write(行數 + 1, 45, 工作表.iloc[行數, 45], 日期格式_置中)
        #     worksheet.write(行數 + 1, 36, 工作表.iloc[行數, 36], 日期格式_靠左)
        #     worksheet.write(行數 + 1, 37, 工作表.iloc[行數, 37], 日期格式_靠左)

    writer.close()


def 文件讀取與輸出(起始日期, 結束日期, 文件路徑):
    日期區間 = 取得日期區間(起始日期, 結束日期)
    # 讀取文件，並分成DIP和SMT
    data_DIP, data_SMT = 文件讀取(文件路徑)

    # 分別對文件進行處理，包括刪除不必要的資料和重複值等
    data_DIP = 資料預處理(data_DIP)
    data_SMT = 資料預處理(data_SMT)

    # 根據使用者指定的日期去抓取文件中的目標足標
    DIP目標項目 = 抓取目標項目(data_DIP, 日期區間)
    SMT目標項目 = 抓取目標項目(data_SMT, 日期區間)

    # 根據目標項目，抓取處理好的文件內的目標
    # 注意data與目標項目必須兩兩相對，比如給data_DIP，那目標項就必須是DIP目標項目
    data_DIP, data_SMT = 目標項目與資料比對(data=[data_DIP, data_SMT], 目標項目=[DIP目標項目, SMT目標項目])

    # 此時需再重置一次足標
    data_DIP = data_DIP.reset_index(drop=True)
    data_SMT = data_SMT.reset_index(drop=True)

    _, _, _, _, _, _, data_DIP = 日期格式與排序的類別標示(日期區間, data_DIP)
    _, _, _, _, _, _, data_SMT = 日期格式與排序的類別標示(日期區間, data_SMT)

    格式更改(日期區間, data=[data_DIP, data_SMT])
