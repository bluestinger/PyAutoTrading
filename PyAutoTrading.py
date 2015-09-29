# -*- encoding: utf8 -*-
__author__ = '人在江湖'

import tkinter.messagebox
from tkinter import *
from tkinter.ttk import *
import time
import win32gui
import win32api
import datetime
import threading

import win32con
import tushare as ts

TIME = 0.1

is_start = False
is_monitor = True
items_lst = []
items_time_lst = []
order_msg = []
stock_code = ''
stock_name = ''
stock_price = ''


def hold_mouse():
    # 固定鼠标位置
    # 获取屏幕的像素大小
    screen_width_pixel = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
    # 固定鼠标位置
    mouse_position = [
        screen_width_pixel - 200, 200, screen_width_pixel - 200, 200]
    win32api.ClipCursor(mouse_position)


def release_mouse():
    # 释放鼠标
    mouse_position = [0, 0, 0, 0]
    win32api.ClipCursor(mouse_position)


def focus_on_window(hwnd):
    # win32gui.ShowWindow(hwnd_parent, win32con.SW_SHOWMAXIMIZED)
    win32gui.SetForegroundWindow(hwnd)


def get_sub_handlers_lst(hwnd, class_name):
    # 查找父窗口下class_name的所有句柄，加为列表
    hwnd_child_lst = []
    hwnd_child = win32gui.FindWindowEx(hwnd, None, class_name, None)
    hwnd_child_lst.append(hwnd_child)
    while True:
        hwnd_child = win32gui.FindWindowEx(hwnd, hwnd_child, class_name, None)
        if hwnd_child != 0:
            hwnd_child_lst.append(hwnd_child)
        else:
            return hwnd_child_lst


def get_last_handlers_lst(hwnd_parent):
    '''
    :param hwnd_parent: 主程序句柄
    :return: 控件句柄列表
    '''
    # 获取双向委托界面dialog下所有控件句柄
    hwnd_second = win32gui.FindWindowEx(hwnd_parent, None, 'AfxMDIFrame42s', None)
    # AfxMDIFrame42s窗口下所以子窗口的句柄
    hwnd_three_lst = get_sub_handlers_lst(hwnd_second, '#32770')
    # 双向委托界面下的dialog类中子窗口句柄，包括EDIT,BUTTON,STATIC控件等
    for handler in hwnd_three_lst:
        hwnd_last_lst = get_sub_handlers_lst(handler, None)
        if len(hwnd_last_lst) == 70:
            # print('hwnd_last_lst are', hwnd_last_lst)
            return hwnd_last_lst


def left_mouse_click(hwnd):
    # 鼠标左键点击
    win32gui.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, None, None)
    time.sleep(TIME)
    win32gui.PostMessage(hwnd, win32con.WM_LBUTTONUP, None, None)
    time.sleep(TIME)


def virtual_key(hwnd_parent, key_code):
    win32gui.PostMessage(hwnd_parent, win32con.WM_KEYDOWN, key_code, 0)  # 消息键盘
    time.sleep(TIME)
    win32gui.PostMessage(hwnd_parent, win32con.WM_KEYUP, key_code, 0)
    time.sleep(TIME)


def input_string(hwnd_edit, string):
    # EDIT控件中输入字符串
    win32api.SendMessage(hwnd_edit, win32con.WM_SETTEXT, None, string)
    time.sleep(TIME)


def is_click_popup_window(hwnd_parent, button_title):
    # 如果有弹出式窗口，点击它的确定按钮
    if hwnd_parent:
        hwnd_popup = win32gui.GetWindow(hwnd_parent, win32con.GW_ENABLEDPOPUP)
        # print('hwnd_parent and hwnd_popup are', hwnd_parent, hwnd_popup)
        if hwnd_popup:
            focus_on_window(hwnd_popup)
            hwnd_button = win32gui.FindWindowEx(hwnd_popup, None, 'Button', button_title)
            left_mouse_click(hwnd_button)
            return True
    return False


def buy(hwnd_parent, stock_code, stock_number):
    virtual_key(hwnd_parent, win32con.VK_F6)  # 必须保证在双向委托界面下才有效
    hwnd_lst = get_last_handlers_lst(hwnd_parent)  # 在买卖前，重新获得句柄
    if is_click_popup_window(hwnd_parent, None):  # 判断是否超时，重新连接
        time.sleep(5)
    left_mouse_click(hwnd_lst[2])
    time.sleep(2)
    input_string(hwnd_lst[2], stock_code)
    left_mouse_click(hwnd_lst[7])
    input_string(hwnd_lst[7], stock_number)
    left_mouse_click(hwnd_lst[8])
    time.sleep(0.5)  # 等0.5s后，再确认窗口有无弹出
    return not is_click_popup_window(hwnd_parent, None)  # 如果返回True，也说明买卖没有出错


def sell(hwnd_parent, stock_code, stock_number):
    virtual_key(hwnd_parent, win32con.VK_F6)
    hwnd_lst = get_last_handlers_lst(hwnd_parent)
    if is_click_popup_window(hwnd_parent, None):
        time.sleep(5)
    left_mouse_click(hwnd_lst[11])
    time.sleep(2)
    input_string(hwnd_lst[11], stock_code)
    left_mouse_click(hwnd_lst[16])
    input_string(hwnd_lst[16], stock_number)
    left_mouse_click(hwnd_lst[17])
    time.sleep(0.5)
    return not is_click_popup_window(hwnd_parent, None)


def order(hwnd_parent, stock_code, stock_number, trade_direction):
    if trade_direction == 'B':
        return buy(hwnd_parent, stock_code, stock_number)
    if trade_direction == 'S':
        return sell(hwnd_parent, stock_code, stock_number)


def trading_init(trading_program_title):
    # 获取交易软件句柄
    hwnd_parent = win32gui.FindWindow(None, trading_program_title)
    if hwnd_parent == 0:
        tkinter.messagebox.showerror('错误', '请先打开华泰证券交易软件，再运行本软件')
    return hwnd_parent


def get_stock_data(stock_code):
    try:
        df = ts.get_realtime_quotes(stock_code)
        return df['name'][0], float(df['price'][0])
    except:
        tkinter.messagebox.showerror('错误', '股票代码错误或网络不稳定，请按确认及停止按钮，检查后重新开始')
        time.sleep(3)
        return '数据错误', -1


def monitor():
    # 股价监控函数
    global stock_name, stock_price, order_msg
    hwnd_parent = trading_init('网上股票交易系统5.0')
    is_traded_first = [True] * 4
    is_traded_second = [True] * 2
    # 如果hwnd_parent为零，直接终止循环
    while is_monitor and hwnd_parent:
        if is_start and stock_code:
            stock_name, stock_price = get_stock_data(stock_code)
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
                        dt_time = dt.time()
                        setting_time = datetime.datetime.strptime(order_time, '%H:%M:%S').time()
                        if dt_time > setting_time:
                            print(dt, setting_time)
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
        self.stock_code_entry = Entry(sub_frame1, textvariable=self.stock_code, width=7)
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

        Combobox(sub_frame2, values=('<', '>'), textvariable=self.lst[0][0], width=2).grid(row=2, column=1, padx=5,
                                                                                           pady=5)
        Spinbox(sub_frame2, from_=0, to=1000, textvariable=self.lst[0][1], increment=0.01, width=6).grid(row=2,
                                                                                                         column=2,
                                                                                                         padx=5, pady=5)
        Combobox(sub_frame2, values=('B', 'S'), textvariable=self.lst[0][2], width=2).grid(row=2, column=3, padx=5,
                                                                                           pady=5)
        Spinbox(sub_frame2, from_=0, to=100000, textvariable=self.lst[0][3], increment=100, width=8).grid(row=2,
                                                                                                          column=4,
                                                                                                          padx=5,
                                                                                                          pady=5)
        Combobox(sub_frame2, values=('<', '>'), textvariable=self.lst[1][0], width=2).grid(row=3, column=1, padx=5,
                                                                                           pady=5)
        Spinbox(sub_frame2, from_=0, to=1000, textvariable=self.lst[1][1], increment=0.01, width=6).grid(row=3,
                                                                                                         column=2,
                                                                                                         padx=5, pady=5)
        Combobox(sub_frame2, values=('B', 'S'), textvariable=self.lst[1][2], width=2).grid(row=3, column=3, padx=5,
                                                                                           pady=5)
        Spinbox(sub_frame2, from_=0, to=100000, textvariable=self.lst[1][3], increment=100, width=8).grid(row=3,
                                                                                                          column=4,
                                                                                                          padx=5,
                                                                                                          pady=5)

        Combobox(sub_frame2, values=('<', '>'), textvariable=self.lst[2][0], width=2).grid(row=4, column=1, padx=5,
                                                                                           pady=5)
        Spinbox(sub_frame2, from_=0, to=1000, textvariable=self.lst[2][1], increment=0.01, width=6).grid(row=4,
                                                                                                         column=2,
                                                                                                         padx=5, pady=5)
        Combobox(sub_frame2, values=('B', 'S'), textvariable=self.lst[2][2], width=2).grid(row=4, column=3, padx=5,
                                                                                           pady=5)
        Spinbox(sub_frame2, from_=0, to=100000, textvariable=self.lst[2][3], increment=100, width=8).grid(row=4,
                                                                                                          column=4,
                                                                                                          padx=5,
                                                                                                          pady=5)

        Combobox(sub_frame2, values=('<', '>'), textvariable=self.lst[3][0], width=2).grid(row=5, column=1, padx=5,
                                                                                           pady=5)
        Spinbox(sub_frame2, from_=0, to=1000, textvariable=self.lst[3][1], increment=0.01, width=6).grid(row=5,
                                                                                                         column=2,
                                                                                                         padx=5, pady=5)
        Combobox(sub_frame2, values=('B', 'S'), textvariable=self.lst[3][2], width=2).grid(row=5, column=3, padx=5,
                                                                                           pady=5)
        Spinbox(sub_frame2, from_=0, to=100000, textvariable=self.lst[3][3], increment=100, width=8).grid(row=5,
                                                                                                          column=4,
                                                                                                          padx=5,
                                                                                                          pady=5)
        # 时间条件单
        sub_frame3 = Frame(frame)
        sub_frame3.pack(padx=5, pady=5)

        Label(sub_frame3, text='19:5:54', width=7).grid(row=1, column=1, padx=5, pady=5)
        Label(sub_frame3, text="方向", width=5).grid(row=1, column=2, padx=5, pady=5)
        Label(sub_frame3, text="数量", width=5).grid(row=1, column=3, padx=5, pady=5)

        self.time_lst = []
        for row in range(2):
            self.time_lst.append([])
            for col in range(3):
                temp = StringVar()
                self.time_lst[row].append(temp)

        Entry(sub_frame3, textvariable=self.time_lst[0][0], width=7).grid(row=2, column=1, padx=5, pady=5)
        Combobox(sub_frame3, values=('B', 'S'), textvariable=self.time_lst[0][1], width=2).grid(row=2, column=2, padx=5,
                                                                                                pady=5)
        Spinbox(sub_frame3, from_=0, to=100000, textvariable=self.time_lst[0][2], increment=100, width=8).grid(row=2,
                                                                                                               column=3,
                                                                                                               padx=5,
                                                                                                               pady=5)
        Entry(sub_frame3, textvariable=self.time_lst[1][0], width=7).grid(row=3, column=1, padx=5, pady=5)
        Combobox(sub_frame3, values=('B', 'S'), textvariable=self.time_lst[1][1], width=2).grid(row=3, column=2, padx=5,
                                                                                                pady=5)
        Spinbox(sub_frame3, from_=0, to=100000, textvariable=self.time_lst[1][2], increment=100, width=8).grid(row=3,
                                                                                                               column=3,
                                                                                                               padx=5,
                                                                                                               pady=5)

        # 日志
        frame2 = Frame(self.window)
        frame2.pack(side=LEFT, padx=10, pady=10)
        scrollbar = Scrollbar(frame2)
        scrollbar.pack(side=RIGHT, fill=Y)
        col_name = ['日期', '时间', '证券代码', '证券名称', '方向', '价格', '数量', '备注']
        self.tree = Treeview(frame2, show='headings', height=10, columns=col_name, yscrollcommand=scrollbar.set)
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
            print(items_time_lst)
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
