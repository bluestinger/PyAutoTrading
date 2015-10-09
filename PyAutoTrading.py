# -*- encoding: utf8 -*-
__author__ = '人在江湖'

import tkinter.messagebox
from tkinter import *
from tkinter.ttk import *
import datetime
import threading
from winguiauto import *
import time
import win32con
import tushare as ts



is_start = False
is_monitor = True
items_lst = []
items_time_lst = []
order_msg = []
stock_code = ''
stock_name = ''
stock_price = ''


def findWantedControls(hwnd):
    # 获取双向委托界面dialog下所有控件句柄
    hwndChildLevel1 = dumpSpecifiedWindow(hwnd, wantedClass='AfxMDIFrame42s')
    hwndChildLevel2 = dumpSpecifiedWindow(hwndChildLevel1[0])
    for handler in hwndChildLevel2:
        hwndChildLevel3 = dumpSpecifiedWindow(handler)
        if len(hwndChildLevel3) == 70:  # 在hwndChildLeve3下，共有70个子窗体
            return hwndChildLevel3


def closePopupWindow(hwnd, wantedText=None, wantedClass=None):
    # 如果有弹出式窗口，点击它的确定按钮
    hwndPopup = findPopupWindow(hwnd)
    if hwndPopup:
        hwndControl = findControl(hwndPopup, wantedText, wantedClass)
        clickButton(hwndControl)
        return True
    return False


def buy(hwnd, stock_code, stock_number):
    pressKey(hwnd, win32con.VK_F6)
    hwndControls = findWantedControls(hwnd)
    # print(hwndControls)
    if closePopupWindow(hwnd, wantedClass='Button'):
        time.sleep(5)
    click(hwndControls[0])
    time.sleep(.5)
    setEditText(hwndControls[0], stock_code)
    time.sleep(.5)
    click(hwndControls[5])
    time.sleep(.5)
    setEditText(hwndControls[5], stock_number)
    time.sleep(.5)
    clickButton(hwndControls[6])
    time.sleep(.5)
    return not closePopupWindow(hwnd, wantedClass='Button')


def sell(hwnd, stock_code, stock_number):
    pressKey(hwnd, win32con.VK_F6)
    hwndControls = findWantedControls(hwnd)
    # print(hwndControls)
    if closePopupWindow(hwnd, wantedClass='Button'):
        time.sleep(5)
    click(hwndControls[7])
    time.sleep(.5)
    setEditText(hwndControls[7], stock_code)
    time.sleep(.5)
    click(hwndControls[12])
    time.sleep(.5)
    setEditText(hwndControls[12], stock_number)
    time.sleep(.5)
    clickButton(hwndControls[13])
    time.sleep(.5)
    return not closePopupWindow(hwnd, wantedClass='Button')


def order(hwnd_parent, stock_code, stock_number, trade_direction):
    if trade_direction == 'B':
        return buy(hwnd_parent, stock_code, stock_number)
    if trade_direction == 'S':
        return sell(hwnd_parent, stock_code, stock_number)


def tradingInit():
    # 获取交易软件句柄
    hwnd_parent = findSpecifiedTopWindow(wantedText='网上股票交易系统5.0')
    if hwnd_parent == 0:
        tkinter.messagebox.showerror('错误', '请先打开华泰证券交易软件，再运行本软件')
        return
    return hwnd_parent


def getStockData(stock_code):
    try:
        df = ts.get_realtime_quotes(stock_code)
        return df['name'][0], float(df['price'][0])
    except:
        return '连接错误', -1


def monitor():
    # 股价监控函数
    global stock_name, stock_price, order_msg
    hwnd_parent = tradingInit()
    is_traded_first = [True] * 4
    is_traded_second = [True] * 2
    # 如果hwnd_parent为零，直接终止循环
    while is_monitor and hwnd_parent:
        if is_start and stock_code:
            stock_name, stock_price = getStockData(stock_code)
            # print(stock_name, stock_price)
            if stock_name != '' and stock_price >= 0:
                # 关系条件单
                index_first = 0
                for relationship, setting_price, direction, number in items_lst:
                    if is_traded_first[index_first] and relationship and direction and \
                                    setting_price != 0 and number != '0':
                        if relationship == '>' and stock_price > setting_price:
                            dt = datetime.datetime.now()
                            if order(hwnd_parent, stock_code, number, direction):
                                order_msg.append(
                                    (dt.strftime('%x'), dt.strftime('%X'), stock_code, stock_name, direction,
                                     stock_price, number, '成功'))
                            else:
                                order_msg.append(
                                    (dt.strftime('%x'), dt.strftime('%X'), stock_code, stock_name, direction,
                                     stock_price, number, '失败'))
                            is_traded_first[index_first] = False
                        if relationship == '<' and stock_price < setting_price:
                            dt = datetime.datetime.now()
                            if order(hwnd_parent, stock_code, number, direction):
                                order_msg.append(
                                    (dt.strftime('%x'), dt.strftime('%X'), stock_code, stock_name, direction,
                                     stock_price, number, '成功'))
                            else:
                                order_msg.append(
                                    (dt.strftime('%x'), dt.strftime('%X'), stock_code, stock_name, direction,
                                     stock_price, number, '失败'))
                            is_traded_first[index_first] = False
                    index_first += 1
                # 时间条件单
                index_second = 0
                for order_time, direction, number in items_time_lst:
                    if is_traded_second[index_second] and order_time and direction and number != '0':
                        print('hello')
                        dt = datetime.datetime.now()
                        current_time = dt.time()
                        set_time = datetime.datetime.strptime(order_time, '%H:%M:%S').time()
                        if current_time > set_time:
                            if order(hwnd_parent, stock_code, number, direction):
                                order_msg.append(
                                    (dt.strftime('%x'), dt.strftime('%X'), stock_code, stock_name, direction,
                                     stock_price, number, '成功'))
                            else:
                                order_msg.append(
                                    (dt.strftime('%x'), dt.strftime('%X'), stock_code, stock_name, direction,
                                     stock_price, number, '失败'))
                            is_traded_second[index_second] = False
                    index_second += 1
        time.sleep(3)


class StockGui:
    def __init__(self):
        self.window = Tk()
        self.window.title("股票交易伴侣")

        # self.window.geometry("800x600+300+300")
        self.window.resizable(0, 0)

        # 股票信息
        frame = Frame(self.window)
        frame.pack(side=LEFT, padx=10, pady=10)

        sub_frame1 = Frame(frame)
        sub_frame1.pack(padx=5, pady=5)

        style = Style()
        style.configure("BW.TLabel", foreground="black", background="red")

        Label(sub_frame1, text="股票代码", width=7).grid(
            row=1, column=1, padx=5, pady=5, sticky=W)
        Label(sub_frame1, text="股票名称", width=7).grid(
            row=1, column=2, padx=5, pady=5, sticky=W)
        Label(sub_frame1, text="当前价格", width=7).grid(
            row=1, column=3, padx=5, pady=5, sticky=W)
        self.stock_code = StringVar()
        self.stock_code_entry = Entry(
            sub_frame1, textvariable=self.stock_code, width=7)
        self.stock_code_entry.grid(row=2, column=1, padx=5, pady=5, sticky=W)
        self.stock_name_label = Label(sub_frame1, width=7, style='BW.TLabel')
        self.stock_name_label.grid(row=2, column=2, padx=5, pady=5, sticky=W)
        self.stock_price_label = Label(sub_frame1, width=7, style='BW.TLabel')
        self.stock_price_label.grid(row=2, column=3, padx=5, pady=5, sticky=W)

        sub_frame2 = Frame(frame)
        sub_frame2.pack(padx=5, pady=5)
        Label(sub_frame2, text="关系", width=5).grid(
            row=1, column=1, padx=5, pady=5, sticky=W)
        Label(sub_frame2, text="价格", width=5).grid(
            row=1, column=2, padx=5, pady=5, sticky=W)
        Label(sub_frame2, text="方向", width=5).grid(
            row=1, column=3, padx=5, pady=5, sticky=W)
        Label(sub_frame2, text="数量", width=5).grid(
            row=1, column=4, padx=5, pady=5, sticky=W)

        self.lst = []
        for row in range(4):
            self.lst.append([])
            for col in range(4):
                temp = StringVar()
                self.lst[row].append(temp)

        Combobox(sub_frame2, values=('<', '>'), textvariable=self.lst[0][0],
                 width=2).grid(row=2, column=1, padx=5, pady=5)
        Spinbox(sub_frame2, from_=0, to=1000, textvariable=self.lst[0][1],
                increment=0.01, width=6).grid(row=2, column=2, padx=5, pady=5)
        Combobox(sub_frame2, values=('B', 'S'), textvariable=self.lst[0][2],
                 width=2).grid(row=2, column=3, padx=5, pady=5)
        Spinbox(sub_frame2, from_=0, to=100000, textvariable=self.lst[0][3],
                increment=100, width=8).grid(row=2, column=4, padx=5, pady=5)
        Combobox(sub_frame2, values=('<', '>'), textvariable=self.lst[1][0],
                 width=2).grid(row=3, column=1, padx=5, pady=5)
        Spinbox(sub_frame2, from_=0, to=1000, textvariable=self.lst[1][1],
                increment=0.01, width=6).grid(row=3, column=2, padx=5, pady=5)
        Combobox(sub_frame2, values=('B', 'S'), textvariable=self.lst[1][2],
                 width=2).grid(row=3, column=3, padx=5, pady=5)
        Spinbox(sub_frame2, from_=0, to=100000, textvariable=self.lst[1][3],
                increment=100, width=8).grid(row=3, column=4, padx=5, pady=5)

        Combobox(sub_frame2, values=('<', '>'), textvariable=self.lst[2][0],
                 width=2).grid(row=4, column=1, padx=5, pady=5)
        Spinbox(sub_frame2, from_=0, to=1000, textvariable=self.lst[2][1],
                increment=0.01, width=6).grid(row=4, column=2, padx=5, pady=5)
        Combobox(sub_frame2, values=('B', 'S'), textvariable=self.lst[2][2],
                 width=2).grid(row=4, column=3, padx=5, pady=5)
        Spinbox(sub_frame2, from_=0, to=100000, textvariable=self.lst[2][3],
                increment=100, width=8).grid(row=4, column=4, padx=5, pady=5)

        Combobox(sub_frame2, values=('<', '>'), textvariable=self.lst[3][0],
                 width=2).grid(row=5, column=1, padx=5, pady=5)
        Spinbox(sub_frame2, from_=0, to=1000, textvariable=self.lst[3][1],
                increment=0.01, width=6).grid(row=5, column=2, padx=5, pady=5)
        Combobox(sub_frame2, values=('B', 'S'), textvariable=self.lst[3][2],
                 width=2).grid(row=5, column=3, padx=5, pady=5)
        Spinbox(sub_frame2, from_=0, to=100000, textvariable=self.lst[3][3],
                increment=100, width=8).grid(row=5, column=4, padx=5, pady=5)
        # 时间条件单
        sub_frame3 = Frame(frame)
        sub_frame3.pack(padx=5, pady=5)

        Label(sub_frame3, text='时:分:秒', width=7).grid(
            row=1, column=1, padx=5, pady=5, sticky=W)
        Label(sub_frame3, text="方向", width=5).grid(
            row=1, column=2, padx=5, pady=5, sticky=W)
        Label(sub_frame3, text="数量", width=5).grid(
            row=1, column=3, padx=5, pady=5, sticky=W)

        self.time_lst = []
        for row in range(2):
            self.time_lst.append([])
            for col in range(3):
                temp = StringVar()
                self.time_lst[row].append(temp)

        Entry(sub_frame3, textvariable=self.time_lst[0][0],
              width=7).grid(row=2, column=1, padx=5, pady=5)
        Combobox(sub_frame3, values=('B', 'S'), textvariable=self.time_lst[0][1],
                 width=2).grid(row=2, column=2, padx=5, pady=5)
        Spinbox(sub_frame3, from_=0, to=100000, textvariable=self.time_lst[0][2],
                increment=100, width=8).grid(row=2, column=3, padx=5, pady=5)
        Entry(sub_frame3, textvariable=self.time_lst[1][0],
              width=7).grid(row=3, column=1, padx=5, pady=5)
        Combobox(sub_frame3, values=('B', 'S'), textvariable=self.time_lst[1][1],
                 width=2).grid(row=3, column=2, padx=5, pady=5)
        Spinbox(sub_frame3, from_=0, to=100000, textvariable=self.time_lst[1][2],
                increment=100, width=8).grid(row=3, column=3, padx=5, pady=5)

        # 日志
        frame2 = Frame(self.window)
        frame2.pack(side=LEFT, padx=10, pady=10)
        scrollbar = Scrollbar(frame2)
        scrollbar.pack(side=RIGHT, fill=Y)
        col_name = ['日期', '时间', '证券代码', '证券名称', '方向', '价格', '数量', '备注']
        self.tree = Treeview(
            frame2, show='headings', height=10, columns=col_name, yscrollcommand=scrollbar.set)
        self.tree.pack()
        scrollbar.config(command=self.tree.yview)
        for name in col_name:
            self.tree.heading(name, text=name)
            self.tree.column(name, width=70, anchor=CENTER)
        #
        # 按钮
        frame3 = Frame(self.window)
        frame3.pack(side=LEFT, padx=10, pady=10)
        self.start_bt = Button(
            frame3, text="开始", command=self.start_stop)
        self.start_bt.pack()
        Button(frame3, text='刷新', command=self.refresh_table).pack()
        self.count = 0

        self.window.protocol(name="WM_DELETE_WINDOW", func=self.close)
        self.window.after(100, self.update_labels)

        self.window.mainloop()

    def refresh_table(self):
        # 刷新机器人委托日志
        length = len(order_msg)
        while self.count < length:
            self.tree.insert('', 0, values=order_msg[self.count])
            self.count += 1

    def update_labels(self):
        # 实时刷新股票价格
        self.stock_name_label['text'] = stock_name
        self.stock_price_label['text'] = str(stock_price)
        self.window.after(3000, self.update_labels)

    def start_stop(self):
        global is_start

        if is_start is False:
            is_start = True
        else:
            is_start = False

        if is_start:
            self.get_items()
            # print(items_time_lst)
            self.start_bt['text'] = '停止'
        else:
            self.start_bt['text'] = '开始'

    def close(self):
        # 关闭软件时，停止monitor线程
        global is_monitor
        is_monitor = False
        self.window.quit()

    def get_items(self):
        global items_lst, stock_code, items_time_lst

        items_lst = []
        items_time_lst = []

        # 获取股票代码
        stock_code = self.stock_code.get().strip()

        # 获取买卖价格数量等
        for row in range(4):
            items_lst.append([])
            for col in range(4):
                temp = self.lst[row][col].get().strip()
                # 把价格列转换浮点数
                if col == 1:
                    items_lst[row].append(float(temp))
                else:
                    items_lst[row].append(temp)

        # 获取时间条件单item
        for row in range(2):
            items_time_lst.append([])
            for col in range(3):
                temp = self.time_lst[row][col].get().strip()
                items_time_lst[row].append(temp)


if __name__ == '__main__':
    t1 = threading.Thread(target=StockGui)
    t2 = threading.Thread(target=monitor)
    t1.start()
    t2.start()
    t1.join()
    t2.join()
