# -*- coding: utf-8 -*-
"""
扯线木偶界面自动化应用编程接口(Puppet UIAutomation API)
技术群：624585416
"""
__author__ = "睿瞳深邃(https://github.com/Raytone-D)"
__project__ = 'Puppet'
__version__ = "0.7.10"
__license__ = 'MIT'

import ctypes
import ctypes.wintypes
import time
import io
import subprocess
import re
from functools import reduce, lru_cache
from collections import OrderedDict

try:
    import pyperclip
except Exception as e:
    print("{}\n请先在命令行下运行：pip install pyperclip，再使用puppet！".format(e))

MSG = {'WM_SETTEXT': 12,
       'WM_GETTEXT': 13,
       'WM_CLOSE': 16,
       'WM_KEYDOWN': 256,
       'WM_KEYUP': 257,
       'WM_COMMAND': 273,
       'BM_CLICK': 245,
       'CB_GETCOUNT': 326,
       'CB_SETCURSEL': 334,
       'CBN_SELCHANGE': 1,
       'COPY': 57634}

user32 = ctypes.windll.user32


class Client:
    """Wrapper for Hexin Tonghuashun Trading Client"""
    NODE = {
        'buy': 161,
        'sell': 162,
        'cancel': 163,
        'cancelable': 163,
        'cancel_order': 163,
        'entrustment': 168,
        'trade': 512,
        'buy2': 512,
        'sell2': 512,
        'account': 512,
        'balance': 512,
        'position': 512,
        'market_value': 165,
        'assets': 165,
        'deals': 512,
        'new': 554,
        'raffle': 554,
        'batch': 5170,
        'bingo': 1070
    }
    MEMBERS = {
        'account': (59392, 0, 1711),
        'mkt': (59392, 0, 1003),
        'balance': (1038,),
        'assets': (1015,),
        'market_value': (1014,),
        'table': (1047, 200, 1047),
        'cancel_order': (3348,),
        'buy': (1003, 1032, 1033, 1018, 1034),
        'sell': (1003, 1032, 1033, 1038, 1034),
        'buy2': (3451, 1032, 1541, 1033, 1018, 1034),
        'sell2': (3453, 1035, 1542, 1058, 1019, 1039)
    }  # 交易市场|证券代码|委托策略|买入价格|可买|买入数量

    ATTRS = ('account', 'balance', 'assets', 'position', 'market_value',
             'entrustment', 'cancelable', 'deals', 'new', 'bingo')
    INIT = 'buy', 'sell', 'cancel', 'trade'
    PAGE = 59648, 59649
    FRESH = 32790
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

    def __init__(self, arg=None):
        """
        :arg: 客户端标题(str)或客户端根句柄(int)
        """
        self.bind(arg)

    "Login"

    def run(self, exe_path):
        assert 'xiadan' in subprocess.os.path.basename(exe_path).split('.')\
            and subprocess.os.path.exists(exe_path), '客户端路径("%s")错误' % exe_path
        print('{} 正在尝试运行客户端("{}")...'.format(time.strftime('%Y-%m-%d %H:%M:%S %a'), exe_path))

        self.pid = subprocess.Popen(exe_path).pid
        pid = ctypes.c_ulong()
        hwndChildAfter = None
        for _ in range(60):
            self.wait(0.5)  # have to
            hLogin = user32.FindWindowExW(None, hwndChildAfter, '#32770', None)  # 用户登录窗口
            user32.GetWindowThreadProcessId(hLogin, ctypes.byref(pid))
            if pid.value == self.pid:
                for _ in range(10):
                    self.wait(0.5)
                    if self.visible(hLogin):
                        self.path = exe_path
                        self.hLogin = hLogin
                        self.root = user32.GetParent(hLogin)
                        break
                break
            else:
                hwndChildAfter = hLogin

        assert hLogin, '客户端没有运行或者找不到用户登录窗口'

    def login(self, account_no=None, password=None, comm_pwd=None, client_path=None, ocr=None, **kwargs):
        """ 重新登录或切换账户
            account_no: 账号, str
            password: 交易密码, str
            comm_pwd: 通讯密码, str
        """
        start = time.time()
        assert client_path, "交易客户端路径不能为空"
        self.run(client_path)
        print('\n{} 正在尝试登入交易服务器...'.format(time.strftime('%Y-%m-%d %H:%M:%S %a')))

        @ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p, ctypes.c_wchar_p)
        def match(handle, args):
            if self.visible(handle):
                user32.GetClassNameW(handle, buf, 32)
                if buf.value == 'Edit':
                    try:
                        text = next(lparam)
                        self.fill(text, handle).wait(0.1)
                    except Exception as e:
                        # print('登录信息填写完毕')
                        return False
            return True

        buf = ctypes.create_unicode_buffer(32)
        if not comm_pwd:
            for _ in range(10):
                self.wait(0.5)
                comm_pwd = self.verify(self.grab(), ocr)
                if len(comm_pwd) >= 4:
                    break
        lparam = [account_no, password, comm_pwd]
        lparam = iter(lparam)
        user32.EnumChildWindows(self.hLogin, match, None)
        self.wait(0.5).click_button(self.hLogin, id_btn=1006)

        for _ in range(30):
            res = self.capture()
            if '暂停登录' in res or '通讯失败' in res:
                raise Exception("%s \n木偶: '交易服务器维护，暂停登录，请稍后再试！'" % res)
            if self.visible():
                print("{} 已登入交易服务器。".format(time.strftime('%Y-%m-%d %H:%M:%S %a')))
                self.birthtime = time.ctime()
                self.acc = account_no
                self.title = self.text(self.root)
                self.elapsed = time.time() - start
                # print('耗时:', self.elapsed )
                self.init()
                self.mkt = (0, 1) if self._text(
                    self.get_handle('mkt')).startswith('上海') else (1, 0)
                return self

    def exit(self):
        "退出系统并关闭程序"
        assert self.visible(), "客户端没有登录"
        user32.PostMessageW(self.root, MSG['WM_CLOSE'], 0, 0)
        print("已退出客户端!")
        return self

    "Trade"

    def trade(self, action, symbol, arg, qty):
        """下单
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
        self.members = iter(self.get_handle(action))
        price = None if isinstance(arg, int) else arg

        self.switch_mkt(symbol).fill(symbol).wait(0.2).switch_way(arg).fill(price)
        full = self._text()
        return self.fill(qty or full).click_button(label={'buy2': '买入[B]',
                                                          'sell2': '卖出[S]'}[action])

    def buy(self, symbol, arg, qty):
        return self.trade('buy2', symbol, arg, qty).answer()

    def sell(self, symbol, arg, qty):
        return self.trade('sell2', symbol, arg, qty).answer()

    def cancel(self, symbol=None, action='cancel_all'):
        return self.cancel_order(symbol, action)

    def cancel_order(self, symbol=None, choice='cancel_all'):
        """ 撤单
            :choice: str, 可选“cancel_buy”、“cancel_sell”或"cancel", "cancel"是撤销指定股票symbol的全部委托。
        """
        self.switch('cancel_order').wait(0.5)  # have to
        editor = self.get_handle('cancel_order')
        if isinstance(symbol, str):
            self.fill(symbol, editor)
            for _ in range(10):
                self.wait(0.3).click_button(label='查询代码')
                hButton = user32.FindWindowExW(self.page, 0, 'Button', '撤单')
                if user32.IsWindowEnabled(hButton):  # 撤单按钮的状态检查
                    break
        return self.click_button(label={
            'cancel': '撤单',
            'cancel_all': '全撤(Z /)',
            'cancel_buy': '撤买(X)',
            'cancel_sell': '撤卖(C)'}[choice]).answer()

    def cancel_all(self):  # 全撤
        return self.cancel_order()

    def cancel_buy(self):  # 撤买
        return self.cancel_order(choice='cancel_buy')

    def cancel_sell(self):  # 撤卖
        return self.cancel_order(choice='cancel_sell')

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

    "Query"

    def query(self, category):
        if category not in self.ATTRS:
            raise AttributeError(category)
        if category in ('account', 'mkt'):
            return self._text(self.get_handle('account'))
        self.switch(category)
        self.click_key({'position': ord('W'),
                        # 'market_value': ord('W'),
                        'deals': ord('E')}.get(category)).wait(0.5)
        if category in ('assets', 'balance', 'market_value'):
            data = float(self._text(self.get_handle(category)))
        else:
            data = self.copy_data(self.get_handle(category))
        # if category is 'market_value':
        #    data = sum(float(pair.get('市值', 0)) for pair in data)
        return data

    def __getattr__(self, attrname):
        return self.query(attrname)

    "Development"

    def __repr__(self):
        return "<%s(ver=%s client=%s root=%s)>" % (
            self.__class__.__name__, __version__, self.client, self.root)

    def bind(self, arg=None):
        """"
        :arg: 客户端的标题或根句柄
        :mkt: 交易市场的索引值
        """
        self.root = arg if isinstance(arg, int) else user32.FindWindowW(0, arg or '网上股票交易系统5.0')
        if self.visible(self.root):
            self.birthtime = time.ctime()
            self.acc = self.account
            self.title = self.text(self.root)
            self.mkt = (0, 1) if self._text(self.get_handle('mkt')).startswith('上海') else (1, 0)
            return self.root

    def visible(self, hwnd=None):
        return user32.IsWindowVisible(hwnd or self.root)

    def switch(self, name):
        assert self.visible(), "客户端已关闭或账户已登出"
        node = name if isinstance(name, int) else self.NODE.get(name)
        if user32.SendMessageW(self.root, MSG['WM_COMMAND'], node, 0):
            self.page = reduce(user32.GetDlgItem, self.PAGE, self.root)
            return self

    def init(self):
        for name in self.INIT:
            self.switch(name).wait(0.3)

        user32.ShowOwnedPopups(self.root, False)
        print('木偶："我准备好了"')
        return self

    def wait(self, timeout=0.5):
        time.sleep(timeout)
        return self

    def copy_data(self, h_table: int):
        "将CVirtualGridCtrl|Custom<n>的数据复制到剪贴板"
        _replace = {'参考市值': '市值', '最新市值': '市值'}  # 兼容国金/平安"最新市值"、银河“参考市值”。
        pyperclip.copy('nan')
        user32.PostMessageW(h_table, MSG['WM_COMMAND'], MSG['COPY'], 0)
        self.wait()

        # 关闭验证码弹窗
        handle = user32.GetLastActivePopup(self.root)
        if handle != self.root:
            for _ in range(9):
                self.wait(0.3)
                if self.visible(handle):
                    text = self.verify(self.grab(handle))
                    hEdit = user32.FindWindowExW(handle, None, 'Edit', "")
                    self.fill(text, hEdit).click_button(handle).wait(0.3)  # have to wait!
                    break

        # 数据格式化
        ret = pyperclip.paste().splitlines()
        temp = (x.split('\t') for x in ret)
        header = next(temp)
        for tag, value in _replace.items():
            if tag in header:
                header.insert(header.index(tag), value)
                header.remove(tag)
        return [OrderedDict(zip(header, x)) for x in temp]

    @lru_cache(None)
    def get_handle(self, action: str):
        """
        :action: 操作标识符
        """
        self.switch(action)
        m = self.MEMBERS.get(action, self.MEMBERS['table'])
        if action in ('buy', 'buy2', 'sell', 'sell2'):
            data = [user32.GetDlgItem(self.page, i) for i in m]
        else:
            data = reduce(user32.GetDlgItem, m, self.root if action in (
                'account', 'mkt') else self.page)
        return data

    def text(self, obj, key=0):
        buf = ctypes.create_unicode_buffer(32)
        {0: user32.GetWindowTextW, 1: user32.GetClassNameW}.get(key)(obj, buf, 32)
        #user32.SendMessageW(obj, MSG['WM_GETTEXT'], 32, buf)
        return buf.value

    def _text(self, h_text=None, id_text=None):
        buf = ctypes.create_unicode_buffer(64)
        if id_text:
            user32.SendDlgItemMessageW(
                self.page, id_text, MSG['WM_GETTEXT'], 64, buf)
        else:
            user32.SendMessageW(
                h_text or next(self.members), MSG['WM_GETTEXT'], 64, buf)
        return buf.value

    def fill(self, text, h_edit=None, h_dialog=None, id_edit=None):
        "fill in"
        h_edit = h_edit or next(self.members)
        if text:
            text = str(text)
            if h_edit:
                user32.SendMessageW(h_edit, MSG['WM_SETTEXT'], 0, text)
            else:
                user32.SendDlgItemMessageW(
                    h_dialog, id_edit, MSG['WM_SETTEXT'], 0, text)
        return self

    def click_button(self, h_dialog=None, label='确定', id_btn=None):
        h_dialog = h_dialog or self.page
        if not id_btn:
            h_btn = user32.FindWindowExW(h_dialog, 0, 'Button', label)
            id_btn = user32.GetDlgCtrlID(h_btn)
        user32.PostMessageW(h_dialog, MSG['WM_COMMAND'], id_btn, 0)
        return self

    def click_key(self, keyCode, param=0):  # 单击按键
        if keyCode:
            user32.PostMessageW(self.page, MSG['WM_KEYDOWN'], keyCode, param)
            user32.PostMessageW(self.page, MSG['WM_KEYUP'], keyCode, param)
        return self

    def grab(self, hParent=None):
        from PIL import ImageGrab

        buf = io.BytesIO()
        rect = ctypes.wintypes.RECT()
        hImage = user32.FindWindowExW(
            hParent or self.hLogin, None, 'Static', "")
        user32.GetWindowRect(hImage, ctypes.byref(rect))
        user32.SetForegroundWindow(hParent or self.hLogin)
        screenshot = ImageGrab.grab((rect.left, rect.top, rect.right + (rect.right - rect.left) * 0.33, rect.bottom))
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

    def capture(self, root=None, label='确定'):
        """ 捕捉弹窗的文本内容 """
        buf = ctypes.create_unicode_buffer(64)
        root = root or self.root
        for _ in range(10):
            self.wait(0.1)
            hPopup = user32.GetLastActivePopup(root)
            if hPopup != root and self.visible(hPopup):
                hTips = user32.FindWindowExW(hPopup, 0, 'Static', None)
                user32.SendMessageW(hTips, MSG['WM_GETTEXT'], 64, buf)
                hButton = user32.FindWindowExW(hPopup, 0, 'Button', label)
                if not hButton:
                    label = '是(&Y)'
                    hButton = user32.FindWindowExW(hPopup, 0, 'Button', label)
                self.click_button(hPopup, label=label)
                break
        text = buf.value
        return text if text else '木偶:"没有回应"'

    def answer(self):
        text = self.capture()
        if any(('小数部分' in text,)):
            print(text)
            text = self.capture()
        return (re.findall(r'(\w*[0-9]+)\w*', text)[0], text) if '合同编号' in text else (0, text)

    def refresh(self):
        user32.PostMessageW(self.root, MSG['WM_COMMAND'], self.FRESH, 0)
        return self if self.visible() else False

    def switch_combo(self, index, hCombo=None):
        handle = hCombo or next(self.members)
        user32.SendMessageW(handle, MSG['CB_SETCURSEL'], index, 0)
        r = user32.SendMessageW(user32.GetParent(handle),
                                MSG['WM_COMMAND'],
                                MSG['CBN_SELCHANGE'] << 16 | user32.GetDlgCtrlID(handle),
                                handle)
        if not r:
            print('切换失败', r)
        return self

    def switch_mkt(self, symbol: str):
        """
        :Prefix:上交所: '5'基, '6'A, '7'申购, '11'转债', 9'B
        适配银河|中山证券的默认值(0 ->上海Ａ股)。注意全角字母Ａ
        """
        index = self.mkt[0] if symbol.startswith(('6', '5', '7', '11')) else self.mkt[1]
        return self.switch_combo(index, next(self.members))

    def switch_way(self, index):
        "切换委托策略"
        return self.switch_combo(index if index in (1, 2, 3, 4, 5) else 0)

    def if_fund(self, symbol, price):
        if symbol.startswith('5'):
            if len(str(price).split('.')[1]) == 3:
                self.capture()

    def summary(self):
        return vars(self)
