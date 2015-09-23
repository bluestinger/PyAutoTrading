# -*- encoding: utf8 -*-
__author__ = '人在江湖'

import tkinter.messagebox
import tkinter as tk
import threading
import time
import win32gui
import win32api
import datetime

import win32con
import tushare as ts

TIME = 100

is_start = False
is_monitor = True
items_list = []
trading_message = []
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
    win32api.Sleep(TIME)
    win32gui.PostMessage(hwnd, win32con.WM_LBUTTONUP, None, None)
    win32api.Sleep(TIME)


def virtual_key(hwnd_parent, key_code):
    win32gui.PostMessage(hwnd_parent, win32con.WM_KEYDOWN, key_code, 0)  # 消息键盘
    win32api.Sleep(TIME)
    win32gui.PostMessage(hwnd_parent, win32con.WM_KEYUP, key_code, 0)
    win32api.Sleep(TIME)


def input_string(hwnd_edit, string):
    # EDIT中输入字符串
    win32api.SendMessage(hwnd_edit, win32con.WM_SETTEXT, None, string)
    win32api.Sleep(TIME)


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
    if is_click_popup_window(hwnd_parent, None):  # 判断是否有重新连接对话框
        time.sleep(5)
    left_mouse_click(hwnd_lst[2])
    input_string(hwnd_lst[2], stock_code)
    left_mouse_click(hwnd_lst[7])
    input_string(hwnd_lst[7], stock_number)
    left_mouse_click(hwnd_lst[8])
    time.sleep(0.5)  # 等0.5s后，在确认有无弹出
    return not is_click_popup_window(hwnd_parent, None)  # 如果返回True，也说明买卖没有出错


def sell(hwnd_parent, stock_code, stock_number):
    virtual_key(hwnd_parent, win32con.VK_F6)
    hwnd_lst = get_last_handlers_lst(hwnd_parent)
    if is_click_popup_window(hwnd_parent, None):
        time.sleep(5)
    left_mouse_click(hwnd_lst[11])
    input_string(hwnd_lst[11], stock_code)
    left_mouse_click(hwnd_lst[16])
    input_string(hwnd_lst[16], stock_number)
    left_mouse_click(hwnd_lst[17])
    time.sleep(0.5)
    return not is_click_popup_window(hwnd_parent, None)


def trading_init(trading_program_title):
    # 获取交易软件句柄
    hwnd_parent = win32gui.FindWindow(None, trading_program_title)
    if hwnd_parent == 0:
        tkinter.messagebox.showerror('错误', '请先打开华泰证券交易软件，再运行本软件')
    return hwnd_parent


def get_stock_data(stock_code):
    stock_data = []
    df = ts.get_realtime_quotes(stock_code)
    # print(df)
    stock_data.append(df['name'][0])
    stock_data.append(float(df['price'][0]))
    return stock_data


def is_digit(str1):
    # 字符串是否是数字
    if str1 == '':
        return False
    for ch in str1:
        if not ch.isdigit():
            if ch != '.':
                return False
    return True


def monitor():
    # 股价监控函数
    global stock_name, stock_price, trading_message
    hwnd_parent = trading_init('网上股票交易系统5.0')
    sell_times = 1
    buy_times = 1
    # 如果hwnd_parent为零，直接终止循环
    while is_monitor and hwnd_parent:
        if is_start:
            if items_list[0] != '':
                stock_name, stock_price = get_stock_data(items_list[0])

                # 卖出
                if items_list[3] != '':

                    # stop_loss_sell
                    if sell_times and (stock_price < items_list[1]):
                        trading_time = datetime.datetime.now()
                        if sell(hwnd_parent, items_list[0], items_list[3]):
                            trading_message.append(
                                (trading_time, '止损', items_list[0], stock_name, items_list[3], stock_price, '成功'))
                        else:
                            trading_message.append(
                                (trading_time, '止损', items_list[0], stock_name, items_list[3], stock_price, '失败'))
                        sell_times -= 1
                        time.sleep(1)

                    # stop_profit_sell
                    if sell_times and (stock_price > items_list[2]):
                        trading_time = datetime.datetime.now()
                        if sell(hwnd_parent, items_list[0], items_list[3]):
                            trading_message.append(
                                (trading_time, '止盈', items_list[0], stock_name, items_list[3], stock_price, '成功'))
                        else:
                            trading_message.append(
                                (trading_time, '止盈', items_list[0], stock_name, items_list[3], stock_price, '失败'))
                        sell_times -= 1
                        time.sleep(1)

                # 卖入
                if items_list[5] != '':
                    # 突破买入
                    if buy_times and (stock_price > items_list[4]):
                        trading_time = datetime.datetime.now()
                        if buy(hwnd_parent, items_list[0], items_list[5]):
                            trading_message.append(
                                (trading_time, '买入', items_list[0], stock_name, items_list[3], stock_price, '成功'))
                        else:
                            trading_message.append(
                                (trading_time, '买入', items_list[0], stock_name, items_list[3], stock_price, '失败'))
                        buy_times -= 1
                        time.sleep(1)
        time.sleep(3)


class StockGui:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("股票交易伴侣")

        frame1 = tk.Frame(self.window)
        frame1.pack(padx=10)

        label_frame = tk.LabelFrame(frame1, text="股票")
        label_frame.pack(side=tk.LEFT, padx=5)

        tk.Label(label_frame, text="股票代码", width=10).grid(
            row=1, column=1, sticky=tk.W)
        tk.Label(label_frame, text="股票名称", width=10).grid(
            row=2, column=1, sticky=tk.W)
        tk.Label(label_frame, text="当前价格", width=10).grid(
            row=3, column=1, sticky=tk.W)
        self.stock_code = tk.StringVar()
        self.stock_code_entry = tk.Entry(label_frame, textvariable=self.stock_code, width=10,
                                         justify=tk.RIGHT)
        self.stock_code_entry.grid(row=1, column=2)
        self.stock_name_label = tk.Label(
            label_frame, width=10, bg="yellow", justify=tk.RIGHT)
        self.stock_name_label.grid(row=2, column=2)
        self.stock_price_label = tk.Label(
            label_frame, width=10, bg="yellow", justify=tk.RIGHT)
        self.stock_price_label.grid(row=3, column=2)

        # 卖出
        label_frame1 = tk.LabelFrame(frame1, text="卖出")
        label_frame1.pack(side=tk.LEFT, padx=5)
        tk.Label(label_frame1, text="止损价格", width=10, fg="blue").grid(
            row=1, column=1, sticky=tk.W)
        tk.Label(label_frame1, text="止盈价格", width=10, fg="blue").grid(
            row=2, column=1, sticky=tk.W)
        tk.Label(label_frame1, text="卖出数量", width=10, fg="blue").grid(
            row=3, column=1, sticky=tk.W)
        self.stop_loss_price = tk.StringVar()
        self.stop_loss_price_entry = tk.Entry(label_frame1, textvariable=self.stop_loss_price, width=10,
                                              justify=tk.RIGHT)
        self.stop_loss_price_entry.grid(row=1, column=2)

        self.stop_profit_price = tk.StringVar()
        self.stop_profit_price_entry = tk.Entry(label_frame1, textvariable=self.stop_profit_price, width=10,
                                                justify=tk.RIGHT)
        self.stop_profit_price_entry.grid(row=2, column=2)

        self.sell_stock_number = tk.StringVar()
        self.sell_stock_number_entry = tk.Entry(label_frame1, textvariable=self.sell_stock_number, width=10,
                                                justify=tk.RIGHT)
        self.sell_stock_number_entry.grid(row=3, column=2)

        # 买入
        ##########################################################
        label_frame2 = tk.LabelFrame(frame1, text="买入")
        label_frame2.pack(side=tk.LEFT, padx=5)
        tk.Label(label_frame2, text="买入价格", width=10, fg="red").grid(
            row=1, column=1, sticky=tk.W)
        tk.Label(label_frame2, text="买入数量", width=10, fg="red").grid(
            row=2, column=1, sticky=tk.W)

        self.buy_stock_price = tk.StringVar()
        self.buy_stock_price_entry = tk.Entry(label_frame2, textvariable=self.buy_stock_price, width=10,
                                              justify=tk.RIGHT)
        self.buy_stock_price_entry.grid(row=1, column=2)

        self.buy_stock_number = tk.StringVar()
        self.buy_stock_number_entry = tk.Entry(label_frame2, textvariable=self.buy_stock_number, width=10,
                                               justify=tk.RIGHT)
        self.buy_stock_number_entry.grid(row=2, column=2)

        label_frame3 = tk.LabelFrame(self.window, text='委托日志')
        label_frame3.pack(side=tk.TOP, padx=10)

        scrollbar = tk.Scrollbar(label_frame3)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.text = tk.Text(label_frame3, width=65, height=10, wrap=tk.WORD, yscrollcommand=scrollbar.set)
        self.text.pack()
        scrollbar.config(command=self.text.yview)

        self.count = 0

        self.start_bt = tk.Button(
            self.window, text="启动", command=self.start_stop)
        self.start_bt.pack()

        self.window.protocol(name="WM_DELETE_WINDOW", func=self.close)
        self.window.after(500, self.update_labels)

        self.window.mainloop()

    def update_labels(self):
        # 实时刷新股票价格
        self.stock_name_label['text'] = stock_name
        self.stock_price_label['text'] = str(stock_price)
        if len(trading_message) > self.count:
            self.text.insert(tk.END, '%s %s %s %s %s %.2f %s\n' % (
                trading_message[-1][0], trading_message[-1][1], trading_message[-1][2],
                trading_message[-1][3], trading_message[-1][4], trading_message[-1][5], trading_message[-1][6]))
            self.count += 1
        self.window.after(3000, self.update_labels)

    def start_stop(self):
        global is_start

        if is_start is False:
            is_start = True
        else:
            is_start = False

        if is_start:
            self.get_items()
            self.start_bt['text'] = '停止'
            self.disable_widget()
        else:
            self.enable_widget()
            self.start_bt['text'] = '启动'

    def close(self):
        # 关闭软件时，停止monitor线程
        global is_monitor
        is_monitor = False
        self.window.quit()

    def enable_widget(self):
        self.stock_code_entry['state'] = tk.NORMAL
        self.stop_loss_price_entry['state'] = tk.NORMAL
        self.stop_profit_price_entry['state'] = tk.NORMAL
        self.sell_stock_number_entry['state'] = tk.NORMAL
        self.buy_stock_price_entry['state'] = tk.NORMAL
        self.buy_stock_number_entry['state'] = tk.NORMAL

    def disable_widget(self):
        self.stock_code_entry['state'] = tk.DISABLED
        self.stop_loss_price_entry['state'] = tk.DISABLED
        self.stop_profit_price_entry['state'] = tk.DISABLED
        self.sell_stock_number_entry['state'] = tk.DISABLED
        self.buy_stock_price_entry['state'] = tk.DISABLED
        self.buy_stock_number_entry['state'] = tk.DISABLED

    def get_items(self):
        global items_list

        items_list = []

        # 股票代码
        stock_code = self.stock_code.get().strip()
        if stock_code.isdigit() and len(stock_code) == 6:
            items_list.append(stock_code)
        else:
            items_list.append('')

        # 止损价格
        stock_loss_price = self.stop_loss_price.get().strip()
        if is_digit(stock_loss_price):
            items_list.append(float(stock_loss_price))
        else:
            items_list.append(0)

        # 不写止盈价，默认为10000止盈
        stop_profit_price = self.stop_profit_price.get().strip()
        if is_digit(stop_profit_price):
            items_list.append(float(stop_profit_price))
        else:
            items_list.append(10000)

        # 卖出股票数量
        sell_stock_number = self.sell_stock_number.get().strip()
        if sell_stock_number.isdigit():
            items_list.append(sell_stock_number)
        else:
            items_list.append('')

        # 不写突破买入价，默认突破10000时买入
        buy_stock_price = self.buy_stock_price.get().strip()
        if is_digit(buy_stock_price):
            items_list.append(float(buy_stock_price))
        else:
            items_list.append(10000)

        # 买入股票数量
        buy_stock_price = self.buy_stock_number.get().strip()
        if buy_stock_price.isdigit():
            items_list.append(buy_stock_price)
        else:
            items_list.append('')


if __name__ == '__main__':
    t1 = threading.Thread(target=StockGui)
    t2 = threading.Thread(target=monitor)
    t1.start()
    t2.start()
    t1.join()
    t2.join()
