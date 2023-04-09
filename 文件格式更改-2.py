import pandas as pd
import pandas.io.formats.excel
import datetime
import xlsxwriter
import openpyxl


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


def 抓取Output足標(data, 抓取日期):
    目標項目 = set()
    for 足標, value in enumerate(data['OUTPUT']):
        if 抓取日期 in value:
            目標項目.add(足標)
    return 目標項目


def 抓取DIP首件足標(data, 抓取日期):
    目標項目 = set()
    足標記數 = 0
    for i in data['DIP首件產出時間/數量']:
        if 抓取日期 in i:
            目標項目.add(足標記數)
        足標記數 += 1
    return 目標項目


def 格式更改(抓取日期, data):
    # 创建一个 ExcelWriter 对象，并将数据框写入 Excel 文件
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
    writer = pd.ExcelWriter(f'{輸出檔名}', engine='xlsxwriter', datetime_format='mm/dd hh:mm')

    # 一些設定
    列寬_10 = [0, 1, 9]
    列寬_5 = [43, 44, 46]
    隱藏行 = ['D', 'F', 'H', 'I', 'L', 'O', 'Q', 'R', 'S',
           'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB',
           'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ',
           'AM', 'AN', 'AO']

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
        # 因為標題行不算，所以要加1
        總筆數 = len(工作表.index) + 1

        工作表.to_excel(writer, index=False, sheet_name=工作表名稱[index])

        # 獲取工作表
        worksheet = writer.sheets[工作表名稱[index]]
        # 設置底色
        # 假日用-淺藍 #b7dee8 、 當日用-淺綠 #92d050 、 明天用-黃色 #ffff00
        # 此部分注意!!! 原本僅用底色格式時，會造成換行格式消失，也就是單元格內若有多行資料就會全部變成一行
        # 因此加入'text_wrap': True處理
        # 經過ChatGPT說明，會有此現象是因為，套用格式時導致換行字元\n被忽略
        # 加入'text_wrap': True確實是一個解法
        淺綠格式 = writer.book.add_format({'bg_color': '#92d050', 'valign': 'vcenter', 'font_size': 10, 'text_wrap': True})
        淺藍格式 = writer.book.add_format({'bg_color': '#b7dee8', 'valign': 'vcenter', 'font_size': 10, 'text_wrap': True})
        黃色格式 = writer.book.add_format({'bg_color': '#ffff00', 'valign': 'vcenter', 'font_size': 10, 'text_wrap': True})
        # 用for迴圈跑過所有Output
        # data.iloc[行數-1, 42] 貌似會直接跳過標題行，從內容的第一格開始，並且足標為0
        # 這邊所抓的足標是從內容的第一行開始，並且足標為0，所以不需要像前面-1，而是一開始的行設定要+1

        # 就目前的邏輯來看，最後一個日期依定會是要填入黃色的日期，故使用list的pop()函數
        黃底日期 = 抓取日期[len(抓取日期) - 1]
        足標串列 = 抓取Output足標(工作表, 黃底日期)
        for 足標 in 足標串列:
            worksheet.write(足標 + 1, 42, 工作表.iloc[足標, 42], 黃色格式)
        足標串列 = 抓取DIP首件足標(工作表, 黃底日期)
        for 足標 in 足標串列:
            worksheet.write(足標 + 1, 41, 工作表.iloc[足標, 41], 黃色格式)

        for index, value in enumerate(抓取日期):
            if value == '6/2':
                足標串列 = 抓取Output足標(工作表, value)
                for 足標 in 足標串列:
                    worksheet.write(足標 + 1, 42, 工作表.iloc[足標, 42], 淺綠格式)
                足標串列 = 抓取DIP首件足標(工作表, value)
                for 足標 in 足標串列:
                    worksheet.write(足標 + 1, 41, 工作表.iloc[足標, 41], 淺綠格式)

        for 剩餘日期 in 抓取日期:
            # '6/2' 到正式版要修改成當天日期_文字格式
            if 剩餘日期 != 黃底日期 and 剩餘日期 != '6/2':
                足標串列 = 抓取Output足標(工作表, 剩餘日期)
                for 足標 in 足標串列:
                    worksheet.write(足標 + 1, 42, 工作表.iloc[足標, 42], 淺藍格式)
                足標串列 = 抓取DIP首件足標(工作表, 剩餘日期)
                for 足標 in 足標串列:
                    worksheet.write(足標 + 1, 41, 工作表.iloc[足標, 41], 淺藍格式)

        # .set_column(0, 0, 10) 用來設置列寬，3個參數分別為，起始列、結束列和列寬
        for 列 in 列寬_10:
            worksheet.set_column(列, 列, 12)
        for 列 in 列寬_5:
            worksheet.set_column(列, 列, 5)
        # DIP首件設置16
        worksheet.set_column(41, 41, 16)
        # Output設置30
        worksheet.set_column(42, 42, 30)
        # 生管備註設置9 因為日期問題 暫時設置
        worksheet.set_column(45, 45, 10)
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

        # formatdict = {'num_format': 'mm/dd'}
        # date_format = writer.book.add_format()
        # worksheet.set_column('AT:AT', 5, date_format)

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

    writer.save()


data_DIP = pd.read_excel('Output.xlsx', sheet_name='DIP')
data_SMT = pd.read_excel('Output.xlsx', sheet_name='SMT')
抓取日期 = ['6/2', '6/3', '6/4', '6/5', '6/6']

格式更改(抓取日期, data=[data_DIP, data_SMT])
