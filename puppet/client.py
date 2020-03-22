# -*- coding: utf-8 -*-
"""
扯线木偶界面自动化应用编程接口(Puppet UIAutomation API)
技术群：624585416
"""
__author__ = "睿瞳深邃(https://github.com/Raytone-D)"
__project__ = 'Puppet'
__version__ = "0.8.20"
__license__ = 'MIT'

import ctypes
import ctypes.wintypes
import time
import io
import subprocess
import re
import random
import threading
import winreg
import os
import csv
import sys

from functools import reduce, lru_cache
from collections import OrderedDict
from importlib import import_module

try:
    import pandas as pd
except Exception as e:
    print(e)

from . import puppet_util as util


MSG = {
    'WM_SETTEXT': 12,
    'WM_GETTEXT': 13,
    'WM_CLOSE': 16,
    'WM_KEYDOWN': 256,
    'WM_KEYUP': 257,
    'WM_COMMAND': 273,
    'BM_CLICK': 245,
    'CB_GETCOUNT': 326,
    'CB_SETCURSEL': 334,
    'CBN_SELCHANGE': 1,
    'COPY': 57634
}

user32 = ctypes.windll.user32

curr_time = lambda : time.strftime('%F %X %a')


def login(accinfos):
    return Client(accinfos)


def get_rect(obj_handle, ext_rate=0):
    rect = ctypes.wintypes.RECT()
    user32.GetWindowRect(obj_handle, ctypes.byref(rect))
    user32.SetForegroundWindow(user32.GetParent(obj_handle))  # have to
    return rect.left, rect.top, rect.right + (
        rect.right - rect.left) * ext_rate, rect.bottom


def grab(rect):
    return import_module('PIL.ImageGrab').grab(rect)


def image_to_string(image, token={
    'appId': '11645803',
    'apiKey': 'RUcxdYj0mnvrohEz6MrEERqz',
    'secretKey': '4zRiYambxQPD1Z5HFh9VOoPXPK9AgBtZ'}):
    if not isinstance(image, bytes):
        buf = import_module('io').BytesIO()
        image.save(buf, 'png')
        image = buf.getvalue()
    return import_module('aip').AipOcr(**token).basicGeneral(image).get(
        'words_result')[0]['words']


def get_text(obj_handle, num=32):
    buf = ctypes.create_unicode_buffer(num)
    user32.SendMessageW(obj_handle, MSG['WM_GETTEXT'], num, buf)
    return buf.value


def lacate_folder(name='Personal'):
    """Personal Recent
    """
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
        r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
    return winreg.QueryValueEx(key, name)[0]  # dir, type


def wait_for_popup():
    buf = ctypes.create_unicode_buffer(64)
    for _ in range(9):
        hwnd = user32.GetForegroundWindow()
        user32.GetWindowTextW(hwnd, buf, 64)
        if buf.value == '另存为' and user32.IsWindowVisible(hwnd):
            break
        time.sleep(0.2)


def simulate_shortcuts(key1, key2=None):
    KEYEVENTF_KEYUP = 2
    scan1 = user32.MapVirtualKeyW(key1, 0)
    user32.keybd_event(key1, scan1, 0, 0)
    if key2:
        scan2 = user32.MapVirtualKeyW(key2, 0)
        user32.keybd_event(key2, scan2, 0, 0)
        user32.keybd_event(key2, scan2, KEYEVENTF_KEYUP, 0)
    user32.keybd_event(key1, scan1, KEYEVENTF_KEYUP, 0)


def export_data(path: str):
    VK_CONTROL = 17
    VK_ALT = 18
    VK_S = 83
    for x in range(99):
        simulate_shortcuts(VK_CONTROL, VK_S)  # 右键保存 Ctrl+S
        wait_for_popup()
        simulate_shortcuts(VK_ALT, VK_S)  # 按钮保存 Alt+S 或 回车键
        for i in range(99):
            time.sleep(0.05)
            try:
                with open(path) as f:
                    string = f.read()
            except Exception:
                continue
            else:
                os.remove(path)
                if string:
                    print(x, i, 'DONE!')
                    return string
                break


class Client:
    """Wrapper for Hexin Tonghuashun Trading Client"""
    NODE = {
        'buy': 161,
        'sell': 162,
        'cancel': 163,
        'cancel_all': 163,
        'cancel_buy': 163,
        'cancel_sell': 163,
        'cancelable': 163,
        'cancel_order': 163,
        'entrustment': 168,
        'trade': 512,
        'buy2': 512,
        'sell2': 512,
        'mkt': 512,
        'account': 165,
        'balance': 165,
        'free_bal': 165,
        'position': 165,
        'market_value': 165,
        'assets': 165,
        'deals': 167,
        'historical_deals': 510,
        'delivery_order': 176,
        'new': 554,
        'raffle': 554,
        'reverse_repo': 717,
        'purchase': 433,
        'redeem': 434,
        'batch': 5170,
        'bingo': 1070
    }

    ATTRS = ('account', 'balance', 'free_bal', 'assets', 'position', 'market_value',
             'entrustment', 'cancelable', 'deals', 'new', 'bingo', 'delivery_order', 'historical_deals')
    INIT = 'position', 'buy', 'sell', 'cancel', 'deals', 'entrustment', 'assets'
    LOGIN = (1011, 1012, 1001, 1003, 1499)
    ACCOUNT = (59392, 0, 1711)
    MKT = (59392, 0, 1003)
    BALANCE = (1012,)
    FREE_BAL = (1016,)
    ASSETS = (1015,)
    MARKET_VALUE = (1014, )
    TABLE = (1047, 200, 1047)
    MONEY = (1308, 200, 1308)
    CANCEL_ORDER = (3348, )
    # symbol, price, max_qty, qty, quote
    BUY = (1032, 1033, 1034, 0, 1018)
    SELL = (1032, 1033, 1034, 0, 1038)
    REVERSE_REPO = (1032, 1033, 1034, 0, 1018)
    BUY2 = (3451, 1032, 1541, 1033, 1018, 1034)
    SELL2 = (3453, 1035, 1542, 1058, 1019, 1039)
    CANCEL = (3348,)
    PURCHASE = (1032, 1034)
    REDEEM = (1032, 1034)
    PAGE = 59648, 59649
    FRESH = 32790
    QUOTE = 1024
    WAY = {
        0: "LIMIT              限价委托 沪深",
        1: "BEST5_OR_CANCEL    最优五档即时成交剩余撤销 沪深",
        2: "BEST5_OR_LIMIT     最优五档即时成交剩余转限价 上海",
        20: "REVERSE_BEST_LIMIT 对方最优价格 深圳",
        3: "FORWARD_BEST       本方最优价格 深圳",
        4: "BEST_OR_CANCEL     即时成交剩余撤销 深圳",
        5: "ALL_OR_CANCEL      全额成交或撤销 深圳"
    }
    buf_length = 32
    client = '同花顺'

    def __init__(self, accinfos={}, enable_heartbeat=True, copy_protection=False,
        **kwargs):
        """
        :arg: 客户端标题(str)或客户端根句柄(int)
        """
        self.root = 0
        if accinfos:
            self.login(**accinfos)
        elif kwargs:
            self.bind(**kwargs)
        self.heartbeat_stamp = time.time()
        self.enable_heartbeat = enable_heartbeat
        self.make_heartbeat()
        self.copy_protection = copy_protection
        dirname = kwargs.get('dirname') or lacate_folder()
        self.filename = f'{dirname}\\table.xls'
        if os.path.isfile(self.filename):
            # print(f'Remove {path}')
            os.remove(self.filename)
        if util.check_input_mode(self.get_handle('buy')[0]) == 'KB':
            try:
                fill = import_module('keyboard').write
            except Exception:
                fill = import_module('pywinauto.keyboard').send_keys
            def func(*args, **kwargs):
                user32.SetForegroundWindow(self._page)
                for text in args:
                    fill(f'{text}\n')
                return self
            self.fill_and_submit = func

    def run(self, exe_path):
        assert 'xiadan' in subprocess.os.path.basename(exe_path).split('.')\
            and subprocess.os.path.exists(
                exe_path), '客户端路径("%s")错误' % exe_path
        print(f'{curr_time()} 正在尝试运行客户端("{exe_path}")...')
        pid = subprocess.Popen(exe_path).pid
        text = ctypes.c_ulong()
        hwndChildAfter = None
        for _ in range(30):
            self.wait()
            login_h = user32.FindWindowExW(None, hwndChildAfter, '#32770', None)
            combobox_h = user32.GetDlgItem(login_h, 1011)  # ComboBox 0x3F3 1011
            if combobox_h:
                user32.GetWindowThreadProcessId(login_h, ctypes.byref(text))
                if text.value == pid:
                    self._page = login_h
                    self.login_h = login_h
                    break
            hwndChildAfter = login_h
        [self.wait() for _ in range(9) if not self.visible(login_h)]
        *self._handles, h1, h2, self._IMG = [user32.GetDlgItem(
            self._page, i) for i in self.LOGIN]
        self._handles.append(h2 if self.visible(h2) else h1)
        self.root = user32.GetParent(login_h)
        print(f'{curr_time()} 登录界面准备就绪。')

    def login(self, account_no: str ='', password: str ='', comm_pwd: str ='',
        client_path: str=''):
        self.run(client_path)
        print(f'{curr_time()} 正在登录交易服务器...')
        self.fill_and_submit(account_no, password, comm_pwd or image_to_string(
            grab(get_rect(self._IMG))))
        # self.capture()
        if self.visible(times=20):
            self.account_no = account_no
            self.password = password
            self.comm_pwd = comm_pwd
            self.mkt = (0, 1) if get_text(self.get_handle('mkt')).startswith(
                '上海') else (1, 0)
            print(f'{curr_time()} 已登入交易服务器。')
            self.init()
            return self

    def exit(self):
        "退出系统并关闭程序"
        assert self.visible(), "客户端没有登录"
        user32.PostMessageW(self.root, MSG['WM_CLOSE'], 0, 0)
        print("已退出客户端!")
        return self

    def fill_and_submit(self, *args, delay=0.1, label=''):
        user32.SetForegroundWindow(self._page)
        for text, handle in zip(args, self._handles):
            self.fill(text, handle)
            if delay:
                for _ in range(9):
                    max_qty = self._text(self._handles[-1])
                    if max_qty not in (''):
                        break
                    self.wait(delay)
        self.wait(0.1)
        self.click_button(label)
        return self

    def trade(self, action: str, symbol: str ='', *args, delay: float =0.1):
        """下单
        客户端->系统->交易设置->默认买入价格(数量): 空、默认卖出价格(数量):空
        :action: str, 交易方式; "buy2"或"sell2", 'margin'
        :symbol: str, 证券代码; 例如'000001', "510500"
        :arg: float or str of float, 判断为委托价格，int判断为委托策略, 例如'3.32', 3.32, 1
        :qty: int or str of int, 委托数量, 例如100, '100'
            委托策略(注意个别券商自定义索引)
            0 LIMIT              限价委托 沪深
            1 BEST5_OR_CANCEL    最优五档即时成交剩余撤销 上海
            2 BEST5_OR_LIMIT     最优五档即时成交剩余转限价 上海
            1 REVERSE_BEST_LIMIT 对方最优价格 深圳
            2 FORWARD_BEST       本方最优价格 深圳
            3 BEST_OR_CANCEL     即时成交剩余撤销 深圳
            4 BEST5_OR_CANCEL    最优五档即时成交剩余撤销 深圳
            5 ALL_OR_CANCEL      全额成交或撤销 深圳
        """
        self.switch(action)
        if action in ('buy', 'sell', 'reverse_repo', 'perchase', 'redeem'):
            self.switch_mkt(symbol, self.get_handle('mkt'))
        self._handles = self.get_handle(action)
        label = {'cancel_all': '全撤(Z /)',
                 'cancel_buy': '撤买(X)',
                 'cancel_sell': '撤卖(C)'}.get(action)
        if action in ('cancel', 'cancel_all', 'cancel_buy', 'cancel_sell'):
            data = self.query('cancelable')
            if not len(data):
                self.wait()
        return self.fill_and_submit(symbol, *args, delay=delay, label=label).wait().answer()
        # return self.fill(qty or full).click_button(label={'buy': '买入[B]','sell': '卖出[S]','reverse_repo': '确定'}[action])

    def buy(self, symbol: str, arg, qty: int) -> tuple:
        return self.trade('buy', symbol, arg, qty)

    def sell(self, symbol: str, arg, qty: int) -> tuple:
        return self.trade('sell', symbol, arg, qty)

    def reverse_repo(self, symbol: str, price: float, qty: int, delay=0.2) -> tuple:
        """逆回购 R-001 SZ '131810'; GC001 SH '204001' """
        return self.trade('reverse_repo', symbol, price, qty, delay=delay)

    def cancel_all(self) -> tuple:  # 全撤
        return self.trade('cancel_all')

    def cancel_buy(self) -> tuple: # 撤买
        return self.trade('cancel_buy')

    def cancel_sell(self) -> tuple: # 撤卖
        return self.trade('cancel_sell')

    def cancel(self, symbol='') -> tuple:
        self.switch('cancel_order').wait(1)  # have to
        editor = self.get_handle('cancel_order')
        if isinstance(symbol, str):
            self.fill(symbol, editor)
            for _ in range(10):
                self.wait(0.3).click_button(label='查询代码')
                hButton = user32.FindWindowExW(self._page, 0, 'Button', '撤单')
                if user32.IsWindowEnabled(hButton):  # 撤单按钮的状态检查
                    break
        return self.click_button(label='撤单').answer()

    def raffle(self):
        "新股申购"
        def func(ipo):
            symbol = ipo.get('新股代码') or ipo.get('证券代码')
            price = ipo.get('申购价格') or ipo.get('发行价格')
            orders = self.entrustment
            had = [order['证券代码'] for order in orders]
            if symbol in had:
                r = (0, '%s 已经申购' % symbol)
            elif symbol not in had:
                r = self.buy(symbol, price, 0)
            else:
                r = (0, '不可预测的申购错误')
            return r

        target = self.new
        if target:
            return [func(ipo) for ipo in target]

    def fund_purchase(self, symbol: str, amount: int):
        """基金申购"""
        return self.trade('purchase', symbol, amount)

    def fund_redeem(self, symbol:str, share: int):
        """基金赎回"""
        return self.trade('redeem', symbol, share)

    def query(self, category):
        """realtime trading data
        2019-5-19 加入数据缓存功能
        2020-2-6 修复 if-elif
        """
        print('Querying {} on-line...'.format(category))
        self.switch(category)

        if category not in self.ATTRS:
            return 'Unknown'
        elif category in ('account', 'mkt'):
            return self._text(self.get_handle('account'))
        elif category in ('assets', 'balance', 'free_bal', 'market_value'):
            for _ in range(99):
                data = self._text(self.get_handle(category))
                if data:
                    return float(data)
                else:
                    self.wait(0.05)
        else:  # data sheet
            if user32.IsIconic(self.root):
                # print('最小化')
                user32.ShowWindow(self.root, 9)
            user32.SetForegroundWindow(self.root)
            [self.wait(0.1) for _ in range(20) if user32.GetForegroundWindow() != self.root]
            string = export_data(self.filename)
            if 'pd' in globals():
                return pd.read_csv(io.StringIO(string), sep='\t', \
                    dtype={'证券代码': str}).dropna(subset=['证券代码'])
            print('没安装 pandas, 返回一个列表')
            g = csv.reader(string.splitlines(), delimiter='\t')
            header = next(g)
            return [dict(zip(header, (x or None for x in gg))) for gg in g]
            # data = list(csv.DictReader(rows, delimiter='\t'))
            # data = self.copy_data(self.get_handle(category))

    def __getattr__(self, attrname):
        return self.query(attrname)

    "Development"

    def __repr__(self):
        return "<%s(ver=%s client=%s root=%s)>" % (
            self.__class__.__name__, __version__, self.client, self.root)

    def bind(self, arg='', **kwargs):
        """"
        :arg: 客户端的标题或根句柄
        :mkt: 交易市场的索引值
        """
        if 'title' in kwargs or isinstance(arg, str):
            self.root = user32.FindWindowW(0, kwargs.get('title') or (
                arg or '网上股票交易系统5.0'))
        elif 'root' in kwargs or isinstance(arg, int):
            self.root = kwargs.get('root') or arg
        if self.visible(self.root):
            self.birthtime = time.ctime()
            self.acc = self.account
            self.title = self.text(self.root)
            self.mkt = (0, 1) if self._text(
                self.get_handle('mkt')).startswith('上海') else (1, 0)
            self.idx = 0
            self.init()
            return self

    def visible(self, hwnd=None, times=0):
        for _ in range(times or 1):
            val = user32.IsWindowVisible(hwnd or self.root)
            if val:
                return True
            elif times > 0:
                self.wait()

    def switch(self, name):
        self.heartbeat_stamp = time.time()
        assert self.visible(), "客户端已关闭或账户已登出"
        node = name if isinstance(name, int) else self.NODE.get(name)
        if user32.SendMessageW(self.root, MSG['WM_COMMAND'], 0x2000<<16|node, 0):
            self._page = reduce(user32.GetDlgItem, self.PAGE, self.root)
            return self

    def init(self):
        for name in self.INIT:
            self.switch(name).wait(0.3)

        user32.ShowOwnedPopups(self.root, False)
        print(f"{curr_time()} 木偶准备就绪！")
        return self

    def wait(self, timeout=0.5):
        time.sleep(timeout)
        return self

    def copy_data(self, h_table: int):
        "将CVirtualGridCtrl|Custom<n>的数据复制到剪贴板"
        # 代码失效，等待移除。
        _replace = {'参考市值': '市值', '最新市值': '市值'}  # 兼容国金/平安"最新市值"、银河“参考市值”。
        pyperclip.copy('')

        for _ in range(9):
            user32.PostMessageW(h_table, MSG['WM_COMMAND'], MSG['COPY'], 0)

            # 关闭验证码弹窗
            if self.copy_protection:
                print('Removing copy protection...')
                self.wait()  # have to
                handle = user32.GetLastActivePopup(self.root)
                if handle != self.root:
                    for _ in range(9):
                        self.wait(0.3)
                        if self.visible(handle):
                            text = self.verify(self.grab(handle))
                            hEdit = user32.FindWindowExW(handle, None, 'Edit', "")
                            self.fill(text, hEdit).click_button(
                                h_dialog=handle).wait(0.3)  # have to wait!
                            break

            self.wait(0.1)
            ret = pyperclip.paste().splitlines()
            if ret:
                break
            self.wait(0.1)

        # 数据格式化
        temp = (x.split('\t') for x in ret)
        header = next(temp)
        for tag, value in _replace.items():
            if tag in header:
                header.insert(header.index(tag), value)
                header.remove(tag)
        return [OrderedDict(zip(header, x)) for x in temp]

    @lru_cache()
    def get_handle(self, action: str):
        """
        :action: 操作标识符
        """
        if action in ('cancel_all', 'cancel_buy', 'cancel_sell'):
            action = 'cancel'
        self.switch(action)
        m = getattr(self, action.upper(), self.TABLE)
        if action in ('buy', 'buy2', 'sell', 'sell2', 'reverse_repo', 'cancel', 'purchase', 'redeem'):
            data = [user32.GetDlgItem(self._page, i) for i in m]
        else:
            data = reduce(
                user32.GetDlgItem, m,
                self.root if action in ('account', 'mkt') else self._page)
        return data

    def text(self, obj, key=0):
        buf = ctypes.create_unicode_buffer(32)
        {
            0: user32.GetWindowTextW,
            1: user32.GetClassNameW
        }.get(key)(obj, buf, 32)
        # user32.SendMessageW(obj, MSG['WM_GETTEXT'], 32, buf)
        return buf.value

    def _text(self, h_text=None, id_text=None):
        buf = ctypes.create_unicode_buffer(64)
        if id_text:
            user32.SendDlgItemMessageW(self._page, id_text, MSG['WM_GETTEXT'],
                                       64, buf)
        else:
            user32.SendMessageW(h_text or next(self.members),
                                MSG['WM_GETTEXT'], 64, buf)
        return buf.value

    def fill(self, text, h_edit=None, h_dialog=None, id_edit=None):
        h_edit = h_edit or next(self.members)
        if text:
            text = str(text)
            if h_edit:
                user32.SendMessageW(h_edit, MSG['WM_SETTEXT'], 0, text)
            else:
                user32.SendDlgItemMessageW(h_dialog, id_edit,
                                           MSG['WM_SETTEXT'], 0, text)
        return self

    def click_button(self, label='', btn_id=1006, h_dialog=None):
        h_dialog = h_dialog or self._page
        if label:
            btn_h = user32.FindWindowExW(h_dialog, 0, 'Button', label)
            btn_id = user32.GetDlgCtrlID(btn_h)
        user32.PostMessageW(h_dialog, MSG['WM_COMMAND'], btn_id, 0)
        return self

    def click_key(self, keyCode, param=0):  # 单击按键
        if keyCode:
            user32.PostMessageW(self._page, MSG['WM_KEYDOWN'], keyCode, param)
            user32.PostMessageW(self._page, MSG['WM_KEYUP'], keyCode, param)
        return self

    def grab(self, hParent=None):
        "屏幕截图"
        from PIL import ImageGrab

        buf = io.BytesIO()
        rect = ctypes.wintypes.RECT()
        hImage = user32.FindWindowExW(hParent or self.hLogin, None, 'Static',
                                      "")
        user32.GetWindowRect(hImage, ctypes.byref(rect))
        user32.SetForegroundWindow(hParent or self.hLogin)
        screenshot = ImageGrab.grab(
            (rect.left, rect.top, rect.right + (rect.right - rect.left) * 0.33,
             rect.bottom))
        screenshot.save(buf, 'png')
        return buf.getvalue()

    def verify(self, image, ocr=None):
        try:
            from aip import AipOcr
        except Exception as e:
            print(e, '\n请在命令行下执行: pip install baidu-aip')

        conf = ocr or {
            'appId': '11645803',
            'apiKey': 'RUcxdYj0mnvrohEz6MrEERqz',
            'secretKey': '4zRiYambxQPD1Z5HFh9VOoPXPK9AgBtZ'
        }

        ocr = AipOcr(**conf)
        try:
            r = ocr.basicGeneral(image).get('words_result')[0]['words']
        except Exception as e:
            print(e, '\n验证码图片无法识别！')
            r = False
        return r

    def capture(self, root=None, label=''):
        """ 捕捉弹窗的文本内容 """
        buf = ctypes.create_unicode_buffer(64)
        root = root or self.root
        for _ in range(99):
            print(':')
            self.wait(0.1)
            hPopup = user32.GetLastActivePopup(root)
            if hPopup != root:  # and self.visible(hPopup):
                hTips = user32.FindWindowExW(hPopup, 0, 'Static', None)
                # print(hex(hPopup).upper(), hex(hTips).upper())
                hwndChildAfter = None
                for _ in range(9):
                    hButton = user32.FindWindowExW(hPopup, hwndChildAfter, 'Button', 0)
                    user32.SendMessageW(hButton, MSG['WM_GETTEXT'], 64, buf)
                    if buf.value in ('是(&Y)', '确定'):
                        label = buf.value
                        break
                    hwndChildAfter = hButton
                user32.SendMessageW(hTips, MSG['WM_GETTEXT'], 64, buf)
                self.click_button(label, h_dialog=hPopup)
                break
        text = buf.value
        return text if text else '委托成功后是否弹出提示对话框：否'

    def answer(self):
        """2020-2-10 修改逻辑确保回报窗口被关闭"""
        for _ in range(3):
            text = self.capture()
            if '无可撤委托' in text or '提交失败' in text:
                return (0, text)
            elif '编号' in text:
                return re.findall(r'(\w*[0-9]+)\w*', text)[0], text
            print(text)
        return 0, '委托(成交)编号无法获取！'

    def refresh(self):
        print('Refreshing page...')
        user32.PostMessageW(self.root, MSG['WM_COMMAND'], self.FRESH, 0)
        return self if self.visible() else False

    def switch_combo(self, hCombo=None):
        handle = hCombo or next(self.members)
        user32.SendMessageW(
            user32.GetParent(handle), MSG['WM_COMMAND'],
            MSG['CBN_SELCHANGE'] << 16 | user32.GetDlgCtrlID(handle), handle)
        return self

    def switch_mkt(self, symbol: str, handle: int):
        """
        :Prefix:上交所: '5'基, '6'A, '7'申购, '11'转债', 9'B
        适配银河|中山证券的默认值(0 ->上海Ａ股)。注意全角字母Ａ
        """
        index = self.mkt[0] if symbol.startswith(
            ('6', '5', '7', '11')) else self.mkt[1]
        user32.SendMessageW(handle, MSG['CB_SETCURSEL'], index, 0)
        return self.switch_combo(handle)

    def switch_way(self, index):
        "切换委托策略"
        handle = next(self.members)
        if index not in (1, 2, 3, 4, 5):
            index = 0
        user32.SendMessageW(handle, MSG['CB_SETCURSEL'], index, 0)
        return self.switch_combo(handle)

    def if_fund(self, symbol, price):
        if symbol.startswith('5'):
            if len(str(price).split('.')[1]) == 3:
                self.capture()

    def summary(self):
        return vars(self)

    def make_heartbeat(self, time_interval=1680):
        """2019-6-6 新增方法制造心跳
        """

        def refresh_page(time_interval):
            while self.enable_heartbeat:
                if not self.visible():
                    print("OFFLINE!")
                stamp = self.heartbeat_stamp
                remainder = time_interval - (time.time() - stamp)
                secs = random.uniform(remainder/2, remainder)
                time.sleep(secs)

                # 若在休眠期间心跳印记没被修改，则刷新页面并修改心跳印记
                if stamp == self.heartbeat_stamp:
                    # print('Making heartbeat...')
                    self.refresh()
                    self.heartbeat_stamp = time.time()

        threading.Thread(
            target=refresh_page,
            kwargs={'time_interval': time_interval},
            name='heartbeat',
            daemon=True).start()

    def quote(self, codes, df_first=True):
        """有bug未修复！ get latest deal price"""
        self.switch('sell')
        code_h, *_, page_h = self.get_handle('sell')
        handle = user32.GetDlgItem(page_h, self.QUOTE)
        names = ['code', 'price']
        if isinstance(codes, str):
            codes = [codes]
        def _quote(code: str) -> float:
            self.fill(code, code_h).wait(0.1)
            for _ in range(5):
                text = get_text(handle)
                if text != '-':
                    return float(text)
                self.wait(0.1)
        data = [(code, _quote(code)) for code in codes]
        if df_first:
            if not hasattr(self, 'pd'):
                self.pd = import_module('pandas')
            data = self.pd.DataFrame(data, columns=names)
        return data
