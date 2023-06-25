from datetime import datetime, timedelta


def 取得日期區間(起始日期, 結束日期):
    起始日期 = datetime.strptime(起始日期, "%Y/%m/%d")
    結束日期 = datetime.strptime(結束日期, "%Y/%m/%d")

    日期區間 = []
    當前日期 = 起始日期

    while 當前日期 <= 結束日期:
        日期區間.append(當前日期.strftime('%#m/%#d'))
        當前日期 += timedelta(days=1)

    print(日期區間)
    return 日期區間


if __name__ == '__main__':
    起始日期 = '2023/06/24'
    結束日期 = '2023/06/30'
    取得日期區間(起始日期, 結束日期)