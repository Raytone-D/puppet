# -*- coding: utf-8 -*-
"""
扯线木偶界面自动化应用编程接口(Puppet UIAutomation API)
技术群：624585416
"""
__author__ = "睿瞳深邃(https://github.com/Raytone-D)"
__project__ = 'Puppet'
__version__ = "0.7.0"
__license__ = 'MIT'

import ctypes
import time
import io
import subprocess
import re
from functools import reduce
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
        'deals': 512,
        'new': 554,
        'raffle':554,
        'batch': 5170,
        'bingo': 1070
    }
    MEMBERS = {
        'account': (59392, 0, 1711),
        'balance': (1038,),
        'table': (1047, 200, 1047),
        'buy2': (3451, 1032, 1541, 1033, 1018, 1034), # 交易市场|证券代码|委托策略|买入价格|可买|买入数量
        'sell2':(3453, 1035, 1542, 1058, 1019, 1039),
        'cancel_order': (3348,)
    }
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
        #self.client = '同花顺'
        self.bind(arg)

    "Login API"

    def run(self, exe_path):
        assert 'xiadan' in subprocess.os.path.basename(exe_path).split('.')\
            and subprocess.os.path.exists(exe_path), '客户端路径("%s")错误' % exe_path
        print('{} 正在尝试运行客户端("{}")...'.format(time.strftime('%Y-%m-%d %H:%M:%S %a'), exe_path))

        self.pid = subprocess.Popen(exe_path).pid
        pid = ctypes.c_ulong()
        dlg = None
        for i in range(60):
            self.wait() # have to
            hLogin = user32.FindWindowExW(dlg, None, '#32770', '用户登录')
            tid = user32.GetWindowThreadProcessId(hLogin, ctypes.byref(pid))
            if pid.value == self.pid:
                for i in range(10):
                    self.wait()
                    if self.visible(hLogin):
                        self.path = exe_path
                        self.hLogin = hLogin
                        self.root = user32.GetParent(hLogin)
                        break
                break
            else:
                dlg = hLogin

        assert hLogin, '客户端没有运行或者找不到用户登录窗口'

    def login(self, account_no=None, password=None, comm_pwd=None, client_path=None, ocr=None, **kwargs):
        """ 重新登录或切换账户
            account_no: 账号, str
            password: 交易密码, str
            comm_pwd: 通讯密码, str
        """
        start = time.time()
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
                    except:
                        #print('登录信息填写完毕')
                        return False
            return True

        buf = ctypes.create_unicode_buffer(32)
        if not comm_pwd:
            for i in range(3):
                self.wait()
                comm_pwd = self.verify(self.grab(), ocr)
                if len(comm_pwd) >= 4:
                    break
        lparam = [account_no, password, comm_pwd]
        assert all(lparam), '用户登录参数不全'
        lparam = iter(lparam)
        user32.EnumChildWindows(self.hLogin, match, None)
        self.wait().click_button(self.hLogin, idButton=1006)

        for i in range(30):
            res = self.capture()
            if '暂停登录' in res or '通讯失败' in res:
                raise Exception("%s \n木偶: '交易服务器维护，暂停登录，请稍后再试！'" % res)
            if self.visible():
                self.birthtime = time.ctime()
                self.acc = account_no
                self.title = self.text(self.root)
                self.elapsed  = time.time() - start
                print("{} 已登入交易服务器。".format(time.strftime('%Y-%m-%d %H:%M:%S %a')))
                # print('耗时:', self.elapsed )
                return self.init()
                break

    def exit(self):
        "退出系统并关闭程序"
        assert self.visible(), "客户端没有登录"
        user32.PostMessageW(self.root, MSG['WM_CLOSE'], 0, 0)
        return self


    "Trade API"

    def trade(self, action, symbol, arg, qty):
        """下单
        :action: 交易方向, str, "buy2"或"sell2", 'margin'
        :symbol: 证券代码, str; 例如'000001'
        :arg: 委托价格('float'或float)或委托策略(0<int<=5), 例如'3.32', 3.32, 1
        :qty: 委托数量, int或"int", 例如100, '100'
            委托策略
            0 LIMIT              限价委托 沪深
            1 BEST5_OR_CANCEL    最优五档即时成交剩余撤销 沪深
            2 BEST5_OR_LIMIT     最优五档即时成交剩余转限价 上海
            2 REVERSE_BEST_LIMIT 对方最优价格 深圳
            3 FORWARD_BEST       本方最优价格 深圳
            4 BEST_OR_CANCEL     即时成交剩余撤销 深圳
            5 ALL_OR_CANCEL      全额成交或撤销 深圳
        """
        label = {'buy2': '买入[B]', 'sell2': '卖出[S]'}.get(action)
        self.members = self.excute(action)
        full = self.switch_mkt(symbol).fill(symbol).wait(.3).switch_way(arg).fill(arg)._text()
        self.fill(qty if qty else full).click_button(label=label).if_fund(symbol, arg)
        return self

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
        editor, hDialog = self.excute('cancel_order')
        self.wait() # have to
        if isinstance(symbol, str):
            self.fill(symbol, editor)
            for i in range(10):
                self.wait(0.3).click_button(hDialog, label='查询代码')
                hButton = user32.FindWindowExW(hDialog, 0, 'Button', '撤单')
                if user32.IsWindowEnabled(hButton): # 撤单按钮的状态检查
                    break
        return self.click_button(hDialog, label={'cancel': '撤单',
                                                 'cancel_all': '全撤(Z /)',
                                                 'cancel_buy': '撤买(X)',
                                                 'cancel_sell': '撤卖(C)'}[choice]).answer()

    def cancel_all(self): # 全撤
        return self.cancel_order()

    def cancel_buy(self): # 撤买
        return self.cancel_order(choice='cancel_buy')

    def cancel_sell(self): # 撤卖
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


    "Query API"

    def query(self, category):
        self.switch(category)
        handle, hDialog = self.excute(category)
        self.click_key(hDialog, {'position': ord('W'), 'deals': ord('E')}.get(category)).wait()
        return self.copy_data(handle)

    @property
    def account(self):
        handle = self.excute('account')
        return self._text(handle)

    @property
    def balance(self):
        self.members = self.excute('balance')
        return float(self._text())

    @property
    def position(self):
        return self.query('position')

    @property
    def market_value(self):
        ret = self.position
        return sum(float(pair.get('市值', 0)) for pair in ret)

    @property
    def deals(self):
        return self.query('deals')

    @property
    def entrustment(self):
        return self.query('entrustment')

    @property
    def cancelable(self):
        return self.query('cancelable')

    @property
    def new(self):
        return self.query('new')

    @property
    def bingo(self):
        print("\n中签信息以券商短信通知为准。")
        return self.query('bingo')


    "Development API"

    def __repr__(self):
        return "<%s(ver=%s client=%s root=%s)>" % (self.__class__.__name__, __version__, self.client, self.root)

    def bind(self, arg=None):
        ":arg: 客户端的标题或根句柄"
        self.root = arg if isinstance(arg, int) else user32.FindWindowW(0, arg or '网上股票交易系统5.0')
        if self.visible(self.root):
            self.birthtime = time.ctime()
            self.acc = self.account
            self.title = self.text(self.root)
            return self.root

    def visible(self, hwnd=None):
        return user32.IsWindowVisible(hwnd or self.root)

    def switch(self, name=None):
        assert self.visible(), "客户端已关闭或账户已登出"
        node = name if isinstance(name, int) else self.NODE[name]
        if user32.SendMessageW(self.root, MSG['WM_COMMAND'], node, 0):
            return self

    def init(self):
        INIT = {'trade': 512}
        for key in INIT.keys():
            self.switch(key).wait(0.3)
        user32.ShowOwnedPopups(self.root, False)
        print('木偶："我准备好了"')
        return self

    def wait(self, delay=0.5):
        time.sleep(delay)
        return self

    def copy_data(self, hCtrl):
        "将CVirtualGridCtrl|Custom<n>的数据复制到剪贴板"
        _replace = {'参考市值': '市值', '最新市值': '市值'}  # 兼容国金/平安"最新市值"、银河“参考市值”。
        pyperclip.copy('nan')
        user32.SendMessageTimeoutW(hCtrl, MSG['WM_COMMAND'], MSG['COPY'], 0, 1, 300)

        handle = user32.GetLastActivePopup(self.root)
        if handle != self.root:
            for i in range(9):
                self.wait(0.3)
                if self.visible(handle):
                    text = self.verify(self.grab(handle))
                    hEdit = user32.FindWindowExW(handle, None, 'Edit', "")
                    self.fill(text, hEdit).click_button(handle).wait(0.3) # have to wait!
                    break

        ret = pyperclip.paste().splitlines()
        temp = (x.split('\t') for x in ret)
        header = next(temp)
        for tag, value in _replace.items():
            if tag in header:
                header.insert(header.index(tag), value)
                header.remove(tag)
        return [OrderedDict(zip(header, x)) for x in temp]

    #@lru_cache(None)
    def excute(self, action, timeout=3):
        """
            :action: str, 操作, 'buy2' or 'sell2'
        """
        start = time.time()
        self.switch(action)
        members = self.MEMBERS.get(action)

        while True:
            self.wait(0.1)
            if time.time() - start > timeout:
                raise Exception('执行操作超时')
            page = reduce(user32.GetDlgItem, self.PAGE, self.root)
            if page:
                if action == 'account':
                    return reduce(user32.GetDlgItem, members, self.root)
                elif members:
                    obj = [user32.GetDlgItem(page, i) for i in members]
                else:
                    obj = [reduce(user32.GetDlgItem, self.MEMBERS['table'], page)]
                obj.append(page)
                return iter(obj)

    def text(self, obj, key=0):
        buf = ctypes.create_unicode_buffer(32)
        {0: user32.GetWindowTextW, 1: user32.GetClassNameW}.get(key)(obj, buf, 32)
        #user32.SendMessageW(obj, MSG['WM_GETTEXT'], 32, buf)
        return buf.value

    def _text(self, handle=None):
        buf = ctypes.create_unicode_buffer(32)
        user32.SendMessageW(handle or next(self.members), MSG['WM_GETTEXT'], 32, buf)
        return buf.value

    def fill(self, text, editor=None, hParent=None, idEditor=None):
        "fill in"
        hEdit = editor or next(self.members)
        if text not in (1, 2, 3, 4, 5):
            text = str(text)
            if hParent and idEditor:
                r = user32.SendDlgItemMessageW(hParent, idEditor, MSG['WM_SETTEXT'], 0, text)
            else:
                r = user32.SendMessageW(hEdit, MSG['WM_SETTEXT'], 0, text)
        return self

    def click_button(self, hDialog=None, label='确定', idButton=None):
        hDialog = hDialog or next(self.members)
        if not idButton:
            hButton = user32.FindWindowExW(hDialog, 0, 'Button', label)
            idButton = user32.GetDlgCtrlID(hButton)
        user32.PostMessageW(hDialog, MSG['WM_COMMAND'], idButton, 0)
        return self

    def click_key(self, handle, keyCode, param=0):   # 单击
        if keyCode:
            x = user32.PostMessageW(handle, MSG['WM_KEYDOWN'], keyCode, param)
            if x:
                user32.PostMessageW(handle, MSG['WM_KEYUP'], keyCode, param)
            else:
                raise Exception("单击按键失败")
        return self

    def grab(self, hParent=None):
        from PIL import ImageGrab

        buf = io.BytesIO()
        rect = ctypes.wintypes.RECT()
        hImage = user32.FindWindowExW(hParent or self.hLogin, None, 'Static', "")
        user32.GetWindowRect(hImage, ctypes.byref(rect))
        user32.SetForegroundWindow(hParent or self.hLogin)
        screenshot = ImageGrab.grab((rect.left, rect.top, rect.right+(rect.right-rect.left)*0.33, rect.bottom))
        screenshot.save(buf, 'png')
        return buf.getvalue()

    def verify(self, image, ocr=None):
        try:
            from aip import AipOcr
        except Exception as e:
            print('e \n请在命令行下执行: pip install baidu-aip')

        conf = ocr or {
            'appId': '11645803',
            'apiKey': 'RUcxdYj0mnvrohEz6MrEERqz',
            'secretKey': '4zRiYambxQPD1Z5HFh9VOoPXPK9AgBtZ'
        }

        ocr = AipOcr(**conf)
        try:
            r = ocr.basicGeneral(image).get('words_result')[0]['words']
        except Exception as e:
            print('e \n验证码图片无法识别！')
            r = False
        return r

    def capture(self, hParent=None, label='确定'):
        """ 捕捉弹窗输出内容 """
        buf = ctypes.create_unicode_buffer(64)
        for n in range(10):
            self.wait(0.1)
            root = hParent or self.root
            hPopup = user32.GetLastActivePopup(root)
            if hPopup != root and self.visible(hPopup):
                hTips = user32.FindWindowExW(hPopup, 0, 'Static', None)
                user32.SendMessageW(hTips, MSG['WM_GETTEXT'], 64, buf)
                hButton = user32.FindWindowExW(hPopup, 0, 'Button', label)
                if not hButton:
                    label = '是(&Y)'
                    hButton = user32.FindWindowExW(hPopup, 0, 'Button', label)
                self.wait().click_button(hPopup, label=label)
                break
        text = buf.value
        return text if text else '木偶:"没有回应"'

    def answer(self):
        text = self.capture()
        return (re.findall(r'(\w*[0-9]+)\w*', text)[0], text) if '合同编号' in text else (0, text)

    def refresh(self):
        user32.PostMessageW(self.root, MSG['WM_COMMAND'], self.FRESH, 0)
        return self if self.visible() else False

    def switch_combo(self, index, hCombo=None):
        handle = hCombo or next(self.members)
        user32.SendMessageW(handle, MSG['CB_SETCURSEL'], index if isinstance(index, int) else 0, 0)
        r = user32.SendMessageW(user32.GetParent(handle), MSG['WM_COMMAND'],
            MSG['CBN_SELCHANGE']<<16|user32.GetDlgCtrlID(handle), handle)
        return self# if r else False

    def switch_mkt(self, symbol):
        """:mkt: 0: 沪A， 1 深A
        :head:'5'沪基金, '6'沪A, '7'沪申购, 9'沪B
        """
        hCombo = next(self.members)
        if not hasattr(self, 'SHA'):
            # 适配银河证券 0: 深A, 1: 沪A，注意全角字母A
            buf = ctypes.create_unicode_buffer(32)
            user32.SendMessageW(hCombo, MSG['WM_GETTEXT'], 32, buf)
            self.SHA, self.SZA = (0, 1) if len(buf.value) == 4 else (1, 0)
        mkt = self.SHA if symbol.startswith(('6', '5', '7')) else self.SZA
        return self.switch_combo(mkt, hCombo)

    def switch_way(self, index):
        "切换委托策略"
        return self.switch_combo(index if index in (1, 2, 3, 4, 5) else 0)

    def if_fund(self, symbol, price):
        if symbol.startswith('5'):
            if len(str(price).split('.')[1]) == 3:
                self.capture()

    def summary(self):
        return vars(self)
