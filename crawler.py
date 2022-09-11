# -*- coding: utf-8 -*-
import requests
from bs4 import BeautifulSoup
import time
import datetime
import urllib.parse
# 視窗化
import tkinter.font as tkFont
import tkinter as tk
import tkinter.ttk as ttk
import os
# 畫圖
import matplotlib.pyplot as py
import matplotlib.ticker as ticker
import matplotlib.font_manager as fm
import matplotlib.backends.tkagg as tkagg
from PIL import Image, ImageTk
from collections import namedtuple

try:
    from json.decoder import JSONDecodeError
except ImportError:
    JSONDecodeError = ValueError
import matplotlib.pyplot as py
import matplotlib.ticker as ticker
import matplotlib.font_manager as fm

# 建立視窗
win = tk.Tk()
win.title("財報狗簡易版")

# 生出物件
f1 = tkFont.Font(size=10, family="微軟正黑體")  # 字型
f2 = tkFont.Font(size=20, family="微軟正黑體")
lb = tk.Label(text="股票代碼:", height=1, width=3, font=f1)
lb1 = tk.Label(text="時間範圍:", height=1, width=10, font=f1)
lb2 = tk.Label(text="至", height=1, width=1, font=f1)
lbyear1 = tk.Label(text="年", height=1, width=1, font=f1)
lbseason1 = tk.Label(text="季", height=1, width=1, font=f1)
lbyear2 = tk.Label(text="年", height=1, width=1, font=f1)
lbseason2 = tk.Label(text="季", height=1, width=1, font=f1)
txtstocknum = tk.Text(height=1, width=2, font=f1)

btnstart = tk.Button(text="查詢", height=1, width=6, font=f1, command=lambda: clickbtnstart())
btnnp = tk.Button(text="每月營收", height=2, width=8, font=f2, command=lambda: clickbtnnp())
btnre = tk.Button(text="每股盈餘", height=2, width=8, font=f2, command=lambda: clickbtneps())
btneps = tk.Button(text="每股淨值", height=2, width=8, font=f2, command=lambda: clickbtnre())
canvas = tk.Canvas(bg='white', height=500, width=650)


def go(event):  # 处理事件，*args表示可变参数
    print(comboxyear1.get())


def go1(event):  # 处理事件，*args表示可变参数
    print(comboxyear2.get())


def go2(event):  # 处理事件，*args表示可变参数
    print(comboxseason1.get())


def go3(event):  # 处理事件，*args表示可变参数
    print(comboxseason2.get())


comvalue = tk.StringVar()  # 窗体自带的文本，新建一个值
comvalue1 = tk.StringVar()
comvalue2 = tk.StringVar()
comvalue3 = tk.StringVar()
comboxyear1 = ttk.Combobox(win, width=3, textvariable=comvalue)  # 初始化
comboxyear1["values"] = ("102", "103", "104", "105", "106", "107", "108")
comboxyear1.current(0)  # 選擇第一個
comboxyear1.bind("<<ComboboxSelected>>", go)  # 绑定事件,(下拉列表框被選中时，绑定go()函数)
comboxyear2 = ttk.Combobox(win, width=3, textvariable=comvalue1)  # 初始化
comboxyear2["values"] = ("102", "103", "104", "105", "106", "107", "108")
comboxyear2.current(0)  # 選擇第一個
comboxyear2.bind("<<ComboboxSelected>>", go1)  # 绑定事件,(下拉列表框被選中时，绑定go()函数)
comboxseason1 = ttk.Combobox(win, width=2, textvariable=comvalue2)  # 初始化
comboxseason1["values"] = ("01", "02", "03", "04")
comboxseason1.current(0)  # 選擇第一個
comboxseason1.bind("<<ComboboxSelected>>", go2)  # 绑定事件,(下拉列表框被選中时，绑定go()函数)
comboxseason2 = ttk.Combobox(win, width=2, textvariable=comvalue3)  # 初始化
comboxseason2["values"] = ("01", "02", "03", "04")
comboxseason2.current(0)  # 選擇第一個
comboxseason2.bind("<<ComboboxSelected>>", go3)  # 绑定事件,(下拉列表框被選中时，绑定go()函数)
# 固定在對的位置
full = tk.NE + tk.SW
lb.grid(row=1, column=0, sticky=full)
lb1.grid(row=1, column=3, columnspan=3)
lb2.grid(row=1, column=8, sticky=full)
lbyear1.grid(row=0, column=7, sticky=full)
lbseason1.grid(row=0, column=10, sticky=full)
lbyear2.grid(row=2, column=7, sticky=full)
lbseason2.grid(row=2, column=10, sticky=full)
txtstocknum.grid(row=1, column=1, sticky=full)
comboxyear1.grid(row=0, column=6)
comboxyear2.grid(row=2, column=6)

comboxseason1.grid(row=0, column=9)
comboxseason2.grid(row=2, column=9)

btnstart.grid(row=1, column=14)
btnnp.grid(row=5, rowspan=2, column=0, columnspan=2)
btnre.grid(row=8, rowspan=2, column=0, columnspan=2)
btneps.grid(row=11, rowspan=2, column=0, columnspan=2)
canvas.grid(row=4, column=2, columnspan=13, rowspan=10)


def clickbtnstart():
    co_id = txtstocknum.get("1.0", tk.END)
    co_id = int(co_id)
    start_year = comboxyear1.get()  # ex.102
    start_season = comboxseason1.get()  # ex.01
    end_year = comboxyear2.get()  # ex.106
    end_season = comboxseason2.get()  # ex.04

    X = {'01': '01', '02': '04', '03': '07', '04': '10'}
    Y = {'01': '03', '02': '06', '03': '09', '04': '12'}
    start_month = X[start_season]  # ex.01
    end_month = Y[end_season]  # ex.12

    months = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']

    def month_intervals1(sy, sm, ey, em):  # 取得起始時間與結束時間的月間隔
        output = [[sy, sm]]
        while (sy, sm) != (ey, em):
            if int(sm) < 12:
                sm = months[int(sm)]
                output.append([sy, sm])
            elif int(sm) == 12:
                sy = str(int(sy) + 1)
                sm = '01'
                output.append([sy, sm])
        return output
    global month_intervals
    month_intervals = month_intervals1(start_year, start_month, end_year, end_month)
    #print('month_intervals:', month_intervals)

    seasons = ['01', '02', '03', '04']

    def season_intervals1(sy, ss, ey, es):  # 取得起始時間與結束時間的季間隔
        output = [[sy, ss]]
        while (sy, ss) != (ey, es):
            if int(ss) < 4:
                ss = seasons[int(ss)]
                output.append([sy, ss])
            elif int(ss) == 4:
                sy = str(int(sy) + 1)
                ss = '01'
                output.append([sy, ss])
        return output
    global season_intervals
    season_intervals = season_intervals1(start_year, start_season, end_year, end_season)
    #print('season_intervals:', season_intervals)

    # 1.每月營收爬蟲代碼
    # 2.每季EPS爬蟲代碼
    # 3.每股淨值爬蟲代碼
    payload = {
        'encodeURIComponent': '1',
        'step': '1',
        'firstin': '1',
        'off': '1',
        'queryName': 'co_id',
        'inpuType': 'co_id',
        'TYPEK': 'all',
        'isnew': 'false',
        'co_id': co_id,
    }

    def crawler(url, user):  # 爬取網站資料
        res = requests.post(url, data=payload)
        soup = BeautifulSoup(res.text, "html.parser")
        raw = soup.find_all('td', {'class': {'even', 'odd'}})
        result = []
        for total in raw:
            for item in total:
                result.append(item)
        result.reverse()
        if user == 'netprofit':
            if x[1] == '04':
                return float(result[0])
            elif x[1] == '03':
                return float(result[1])
            elif x[1] == '02':
                return float(result[2])
            else:
                return float(result[3])
        elif user == 'eps':
            if x[1] == '04':
                return round((float(result[0]) - float(result[1])), 2)
            elif x[1] == '03':
                return round((float(result[1]) - float(result[2])), 2)
            elif x[1] == '02':
                return round((float(result[2]) - float(result[3])), 2)
            else:
                return round(float(result[3]), 2)
    global oprev_list
    oprev_list = []
    for x in month_intervals:
        payload['year'], payload['month'] = x[0], x[1]
        res = requests.post("http://mops.twse.com.tw/mops/web/t05st10_ifrs", data=payload)
        soup = BeautifulSoup(res.text, "html.parser")
        oprev = soup.find('td', {'class': 'odd'})
        oprev_list.append(int(oprev.get_text().strip().replace(',', '')))
        if len(oprev_list) % 20 == 0:  # 若超過20個月不休息會被ban掉(網站有反爬蟲)
            time.sleep(60)
    print('oprev_list:', oprev_list)
    global netprofit_list
    global eps_list
    netprofit_list = []
    eps_list = []
    for x in season_intervals:
        payload['year'], payload['season'] = x[0], x[1]
        netprofit = crawler("http://mops.twse.com.tw/mops/web/t163sb16", 'netprofit')
        eps = crawler("http://mops.twse.com.tw/mops/web/t163sb15", 'eps')
        netprofit_list.append(netprofit)
        netprofit_list.append(0)
        netprofit_list.append(0)
        eps_list.append(eps)
        eps_list.append(0)
        eps_list.append(0)
        if len(netprofit_list) % 20 == 0 or len(eps_list) % 20 == 0:  # 若超過20個月不休息會被ban掉(網站有反爬蟲)
            time.sleep(60)
            
    print('netprofit_list:', netprofit_list)
    print('eps_list:', eps_list)

    # 4.股價爬蟲代碼

    start_year = int(month_intervals[0][0]) + 1911
    start_month = int(month_intervals[0][1])
    end_year = int(month_intervals[-1][0]) + 1911
    end_month = int(month_intervals[-1][1])

    TWSE_BASE_URL = 'http://www.twse.com.tw/'
    DATATUPLE = namedtuple('Data', ['date', 'capacity', 'turnover', 'open',
                                    'high', 'low', 'close', 'change', 'transaction'])

    class TWSEFetcher():

        REPORT_URL = urllib.parse.urljoin(TWSE_BASE_URL, 'exchangeReport/STOCK_DAY')

        # 爬取網站資料
        def fetch(self, year: int, month: int, sid: str, retry: int = 5):
            params = {'date': '%d%02d01' % (year, month), 'stockNo': sid}
            for retry_i in range(retry):
                r = requests.get(self.REPORT_URL, params=params)
                try:
                    data = r.json()
                except JSONDecodeError:
                    continue
                else:
                    break
            

            data['data'] = self.purify(data)
            return data
            #print(data)

        # 建立datatuple 轉換成數字格式
        def _make_datatuple(self, data):
            data[0] = datetime.datetime.strptime(self._convert_date(data[0]), '%Y/%m/%d')
            data[1] = int(data[1].replace(',', ''))
            data[2] = int(data[2].replace(',', ''))
            data[3] = float(data[3].replace(',', ''))
            data[4] = float(data[4].replace(',', ''))
            data[5] = float(data[5].replace(',', ''))
            data[6] = float(data[6].replace(',', ''))
            data[7] = float(0.0 if data[7].replace(',', '') == 'X0.00' else data[7].replace(',', ''))
            data[8] = int(data[8].replace(',', ''))
            return DATATUPLE(*data)

        # 去掉資料上方標題列
        def purify(self, original_data):
            #print(type(original_data), original_data)
            return [self._make_datatuple(d) for d in original_data['data']]

        def _convert_date(self, date):
            """Convert '106/05/01' to '2017/05/01'"""
            return '/'.join([str(int(date.split('/')[0]) + 1911)] + date.split('/')[1:])

    class Stock():

        def __init__(self, sid: str, initial_fetch: bool = True):
            self.sid = sid
            self.fetcher = TWSEFetcher()
            self.raw_data = []
            self.data = []
            # Init data
            if initial_fetch:
                self.fetch_from(start_year, start_month, end_year, end_month)

        def _month_year_iter(self, start_month, start_year, end_month, end_year):
            ym_start = 12 * start_year + start_month - 1
            ym_end = 12 * end_year + end_month
            for ym in range(ym_start, ym_end):
                y, m = divmod(ym, 12)
                yield y, m + 1

        def fetch(self, year: int, month: int):
            """Fetch year month data"""
            self.raw_data = [self.fetcher.fetch(year, month, self.sid)]
            self.data = self.raw_data[0]['data']
            return self.data

        def fetch_from(self, start_year: int, start_month: int, end_year: int, end_month: int):
            """Fetch data from year, month to current year month data"""
            self.raw_data = []
            self.data = []
            for year, month in self._month_year_iter(start_month, start_year, end_month, end_year):
                self.raw_data.append(self.fetcher.fetch(year, month, self.sid))
                self.data.extend(self.raw_data[-1]['data'])

            return self.data

        @property
        def date(self):
            return [d.date for d in self.data]

        @property
        def price(self):
            return [d.close for d in self.data]

    stock = Stock(co_id)
    # 計算月平均股價
    prices_list, days_list = [], []
    price, day = 0, 1
    for i in range(len(stock.date)):
        price += float(stock.price[i])
        try:
            if stock.date[i + 1].month == stock.date[i].month:  # 如果明天還是同個月份
                day += 1
            else:
                prices_list.append(price)
                days_list.append(day)
                price, day = 0, 1
        except:
            prices_list.append(price)
            days_list.append(day)
    global stockprice_list
    stockprice_list = []
    for i, j in zip(prices_list, days_list):
        stockprice_list.append(round(i / j, 2))

    print("stockprice_list:", stockprice_list)

def xlabel(xlist1):
    start = int(xlist1[0][0]) + 1911
    end = 2018
    return range(start - 1, end)


font = fm.FontProperties(fname='C:\\Windows\\Fonts\\mingliu.ttc')

def figure(ylist1, name1):
    # 長條圖
    plot1 = py.figure().add_subplot(111)
    if name1 == '每月營收':
        py.bar(range(len(month_intervals)), ylist1, color='#FFD700', width=0.8, edgecolor='#DAA520', label=name1,
               alpha=0.45, align='edge')
    else:
        py.bar(range(len(month_intervals)), ylist1, color='#FFD700', width=1.6, edgecolor='#DAA520', label=name1,
               alpha=0.45, align='edge')
    py.ylim(0, max(ylist1) * 1.2)
    # 折線圖
    ylist2, name2 = stockprice_list, '月均價'
    plot2 = plot1.twinx()
    plot2.plot(range(len(month_intervals)), ylist2, color='#8B0000', label=name2, alpha=0.45)
    # 圖例
    plot2.legend(loc=2, bbox_to_anchor=(0.22, 0.95), prop=font, frameon=False)
    plot1.legend(loc=2, bbox_to_anchor=(0, 0.95), prop=font, frameon=False)
    # x軸刻度
    plot2.xaxis.set_major_locator(ticker.MultipleLocator(12))
    plot2.set_xticklabels(xlabel(month_intervals))

    #canvas = figure(eps_list, '單季EPS')  # 單季EPS + 股價
    #canvas = figure(netprofit_list, '每股淨值')  # 每股淨值 + 股價
def clickbtnnp():
    figure(oprev_list, '每月營收')
    py.savefig('tmp.gif')
    py.show()
    i = ImageTk.PhotoImage(Image.open('tmp.gif'))
    canvas.create_image(0, 0, image=i, anchor= tk.NW)
    #canvas.show()
def clickbtnre():
    figure(netprofit_list, '每股淨值')
    py.savefig('tmp.gif')
    i = ImageTk.PhotoImage(Image.open('tmp.gif'))
    canvas.create_image(0, 0, image=i, anchor= tk.NW)
    #canvas.show()
def clickbtneps():
    figure(eps_list, '單季EPS')
    py.savefig('tmp.gif')
    i = ImageTk.PhotoImage(Image.open('tmp.gif'))
    canvas.create_image(0, 0, image=i, anchor= tk.NW)
    #canvas.show()
win.mainloop()

