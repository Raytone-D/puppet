# -*- coding: utf-8 -*-
"""
扯线木偶界面自动化应用编程接口(Puppet UIAutomation API)
技术群：624585416
"""
__author__ = "睿瞳深邃(https://github.com/Raytone-D)"
__project__ = 'Puppet'
__version__ = "0.6.3"
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

#        'COMBO': (59392, 0, 2322),

TWO_WAY = {'买入代码': 1032,
           '买入价格': 1033,
           '买入数量': 1034,
           '买入': 1006,
           '卖出代码': 1035,
           '卖出价格': 1058,
           '卖出数量': 1039,
           '卖出': 1008,
           '可用余额': 1038,
           '刷新': 32790,
           '报表': 1047}

NEW = {'新股代码': 1032,
       '新股名称': 1036,
       '申购价格': 1033,
       '可申购数量': 1018,
       '申购数量': 1034,
       '申购': 1006}

RAFFLE = ['新股代码', '证券代码', '申购价格']# , '申购上限']

VKCODE = {'F1': 112,
          'F2': 113,
          'F3': 114,
          'F4': 115,
          'F5': 116,
          'F6': 117}

user32 = ctypes.windll.user32


class Puppet:

    NODE = {
        #'buy': 161,
        #'sell': 162,
        'cancel_order': 163,
        'entrustment': 168,
        'raffle':554,
        'trade': 512
    }
    MEMBERS = {
        'account': (59392, 0, 1711),
        'balance': (1038,),
        'table': (1047, 200, 1047),
        'buy2': (3451, 1032, 1033, 1034),
        'sell2':(3453, 1035, 1058, 1039),
        'cancel_order': (3348,)
    }
    PAGE = 59648, 59649

    buf_length = 32

    def __init__(self, title='网上股票交易系统5.0', main=None):
        """
            :title: str, 客户端标题
            :main: int, 客户端主窗口句柄
        """
        print('Puppet TraderApi, version {}\n'.format(__version__))

        self.start = time.time()
        self.root = main or self.get_main(title)
        #self.init()

    def get_main(self, title):
        if isinstance(title, str):
            return user32.FindWindowW(0, title)

    def init(self, **kwargs):
        if self.visible():
            print('木偶："正在热身，请稍候..."\n')
            for key in self.NODE.keys():
                self.switch(key).wait(0.3)

            user32.ShowOwnedPopups(self.root, False)
            print('木偶："我准备好了"\n')
            print("cost: %s\n" % (time.time() - self.start))
        return self

    def run(self, exe_path):
        assert 'xiadan' in subprocess.os.path.basename(exe_path).split('.')[0] and subprocess.os.path.exists(exe_path), '客户端路径错误'
        print('{} 正在尝试运行客户端({})...'.format(time.strftime('%Y-%m-%d %H:%M:%S %a'), exe_path))

        self.pid = subprocess.Popen(exe_path).pid
        pid = ctypes.c_ulong()
        dlg = None
        for i in range(60):
            self.wait() # important
            hLogin = user32.FindWindowExW(dlg, None, '#32770', '用户登录')
            tid = user32.GetWindowThreadProcessId(hLogin, ctypes.byref(pid))
            if pid.value == self.pid and self.visible(hLogin):
                break
            else:
                dlg = hLogin

        assert hLogin, '客户端没有运行或者找不到用户登录窗口'

        self.root = user32.GetParent(hLogin)
        self.hLogin = hLogin
        self.title = self.text(self.root)
        self.path = exe_path

        print("cost:", time.time() - self.start)

    def text(self, obj, key=0):
        buf = ctypes.create_unicode_buffer(32)
        {0: user32.GetWindowTextW, 1: user32.GetClassNameW}.get(key)(obj, buf, 32)
        #user32.SendMessageW(obj, MSG['WM_GETTEXT'], 32, buf)
        return buf.value

    def login(self, account_no=None, password=None, comm_pwd=None, client_path=None, ocr=None, **kwargs):
        """ 重新登录或切换账户
            account_no: 账号, str
            password: 交易密码, str
            comm_pwd: 通讯密码, str
        """
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
            comm_pwd = self.verify(self.grab(), ocr)
        lparam = [account_no, password, comm_pwd]
        assert all(lparam), '用户登录参数不全'
        lparam = iter(lparam)
        user32.EnumChildWindows(self.hLogin, match, None)
        self.wait().click_button(self.hLogin, idButton=1006)

        for i in range(60):
            self.wait()
            if self.visible():
                print("{} 已登入交易服务器。".format(time.strftime('%Y-%m-%d %H:%M:%S %a')))
                print('cost:', time.time() - self.start)
                break

        self.init()
        return self

    def wait(self, delay=0.5):
        time.sleep(delay)
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

        client = AipOcr(**conf)
        try:
            return client.basicGeneral(image).get('words_result')[0]['words']
        except Exception as e:
            raise Exception('e \n验证码图片无法识别！')

    def exit(self):
        "退出系统并关闭程序"
        assert self.visible(), "客户端没有登录"
        user32.PostMessageW(self.root, MSG['WM_CLOSE'], 0, 0)
        return self

    def switch(self, name=None): #root=None):
        node = {
            'buy2': 512,
            'sell2': 512,
            'account': 512,
            'balance': 512,
            'position': 512,
            'deals': 512,
            'new': 554,
            'bingo': 1070,
            'cancelable': 163
        }.get(name) or self.NODE.get(name)
        #print('page', name, node)
        #if isinstance(root, int):
        #    self.root = root
        assert self.visible(), "客户端已关闭或账户已登出"
        assert user32.SendMessageW(self.root, MSG['WM_COMMAND'], node, 0)
        return self

    def click_button(self, hDialog=None, label='确定', idButton=None):
        hDialog = hDialog or next(self.members)
        if not idButton:
            handle = user32.FindWindowExW(hDialog, 0, 'Button', label)
            idButton = user32.GetDlgCtrlID(handle)
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

    def query(self, category):
        self.switch(category)
        handle, hDialog = self.excute(category)
        self.click_key(hDialog, {'position': ord('W'), 'deals': ord('E')}.get(category)).wait()
        return self.copy_data(handle)

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

    def _wait(self, container, id_item):
        buf = ctypes.create_unicode_buffer(32)
        buf.value = ''  # False，待假成真
        for n in range(500):
            time.sleep(0.01)
            user32.SendDlgItemMessageW(container, id_item, MSG['WM_GETTEXT'], 64, buf)
            if buf.value:
                break
        return self

    #@lru_cache(None)
    def excute(self, action, timeout=3):
        """
            :action: str, 操作, 'buy2' or 'sell2'
        """
        self.switch(action)
        members = self.MEMBERS.get(action)
        start = time.time()

        while True:
            self.wait(0.1)
            if time.time() - start > timeout:
                raise Exception('执行操作超时')
            page = reduce(user32.GetDlgItem, self.PAGE, self.root)
            if page:
                if action == 'account':
                    obj = reduce(user32.GetDlgItem, members, self.root)
                elif members:
                    obj = [user32.GetDlgItem(page, i) for i in members]
                else:
                    obj = [reduce(user32.GetDlgItem, self.MEMBERS['table'], page)]
                obj.append(page)
                return iter(obj)

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
        return buf.value

    def answer(self):
        text = self.capture()
        return (re.findall('(\w*[0-9]+)\w*', text)[0], text) if '合同编号' in text else (0, text)

    def trade(self,*args):
        """下单"""
        symbol, price, qty, action, way, *other = args
        self.members = self.excute(action)
        label = {'buy2': '买入[B]', 'sell2': '卖出[S]'}.get(action)
        self.switch_mkt(symbol).fill(symbol).wait(.3).fill(price).fill(qty).wait(0.2).click_button(label=label).if_fund(price)
        return self

    def buy(self, symbol, price, qty, action='buy2', way='limited'):
        """
            :action: str, 'buy', 'buy2'
            :way: str, 'limited', or 'market'
        """
        return self.trade(symbol, price, qty, action, way).answer()

    def sell(self, symbol, price, qty, action='sell2', way='limited'):
        """
            :action: str, 'sell', 'sell2'
            :way: str, 'limited' or 'market'
        """
        return self.trade(symbol, price, qty, action, way).answer()

    def cancel(self, symbol, action='cancel_all'):
        return self.cancel_order(symbol, action)

    def cancel_order(self, symbol=None, choice='cancel_all'):
        """ 撤单
            :choice: str, 可选“cancel_buy”、“cancel_sell”或"cancel", "cancel"是撤销指定股票symbol的全部委托。
        """
        editor, hDialog = self.excute('cancel_order')
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

    def _text(self, handle=None):
        buf = ctypes.create_unicode_buffer(32)
        user32.SendMessageW(handle or next(self.members), MSG['WM_GETTEXT'], 32, buf)
        return buf.value

    @property
    def account(self):
        handle = reduce(user32.GetDlgItem, self.MEMBERS['account'], self.root)
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
        print("\n该券商的中签查询未经测试，注意复查。")
        return self.query('bingo')

    def cancel_all(self): # 全撤
        self.cancel_order()

    def cancel_buy(self): # 撤买
        self.cancel_order(choice='cancel_buy')

    def cancel_sell(self): # 撤卖
        self.cancel_order(choice='cancel_sell')

    def raffle(self, skip=False):    # 打新
        #user32.SendMessageW(self.root, MSG['WM_COMMAND'], NODE['新股申购'], 0)
        #self._raffle = reduce(user32.GetDlgItem, NODE['FORM'], self.root)
        #close_pop()    # 弹窗无需关闭，不影响交易。
        #schedule = self.query(self._raffle)
        buf = ctypes.create_unicode_buffer(32)
        ret = self.new
        if not ret:
            print("是日无新!")
            return ret
        self._raffle = reduce(user32.GetDlgItem, self.MEMBERS['table'], self._container['raffle'])
        self._raffle_parts = {k: user32.GetDlgItem(self._raffle, v) for k, v in NEW.items()}
            #new = [x.split() for x in schedule.splitlines()]
            #index = [new[0].index(x) for x in RAFFLE if x in new[0]]    # 索引映射：代码0, 价格1, 数量2
            #new = map(lambda x: [x[y] for y in index], new[1:])
        for new in ret:
            symbol, price = [new[y] for y in RAFFLE if y in new.keys()]
            if symbol[0] == '3' and skip:
                print("跳过创业板新股: {}".format(symbol))
                continue
            user32.SendMessageW(self._raffle_parts['新股代码'], MSG['WM_SETTEXT'], 0, symbol)
            time.sleep(0.3)
            user32.SendMessageW(self._raffle_parts['申购价格'], MSG['WM_SETTEXT'], 0, price)
            time.sleep(0.3)
            user32.SendMessageW(self._raffle_parts['可申购数量'], MSG['WM_GETTEXT'], 32, buf)
            if not int(buf.value):
                print('跳过零数量新股：{}'.format(symbol))
                continue
            user32.SendMessageW(self._raffle_parts['申购数量'], MSG['WM_SETTEXT'], 0, buf.value)
            time.sleep(0.3)
            user32.PostMessageW(self._raffle, MSG['WM_COMMAND'], NEW['申购'], 0)

        return [new for new in self.cancelable if '配售申购' in new['操作']]

    def switch_combo(self, index, idCombo, hCombo):
        user32.SendMessageW(hCombo, MSG['CB_SETCURSEL'], index, 0)
        user32.SendMessageW(user32.GetParent(hCombo), MSG['WM_COMMAND'], MSG['CBN_SELCHANGE']<<16|idCombo, hCombo)

    def fill(self, text, editor=None, hParent=None, idEditor=None):
        "fill in"
        if hParent and idEditor:
            user32.SendDlgItemMessageW(hParent, idEditor, MSG['WM_SETTEXT'], 0, text)
        else:
            user32.SendMessageW(editor or next(self.members), MSG['WM_SETTEXT'], 0, text)
        return self

    def visible(self, hwnd=None):
        return user32.IsWindowVisible(hwnd or self.root)

    def refresh(self):
        user32.PostMessageW(self._container['trade'], MSG['WM_COMMAND'], TWO_WAY['刷新'], 0)

    def switch_mkt(self, symbol):
        "0: SH, 1: SZ; '5'沪基金, '6'沪A, '9'沪B"
        handle = next(self.members)
        index = 0 if symbol.startswith(('6', '5')) else 1
        user32.SendMessageW(handle, MSG['CB_SETCURSEL'], index, 0)
        user32.SendMessageW(user32.GetParent(handle), MSG['WM_COMMAND'], MSG['CBN_SELCHANGE']<<16|user32.GetDlgCtrlID(handle), handle)
        return self

    def if_fund(self, symbol):
        if len(symbol.split('.')[1]) == 3:
            self.capture()

    def summary(self):
        return vars(self)


if __name__ == '__main__':
    import platform
    print('\n{}\nPython Version: {}'.format(platform.platform(), platform.python_version()))
    gf = {
        'account_no': '666622xxxxxx',
        'password': '123456',
        'comm_pwd': '',
        'exe_path': r'D:\Utils\gfwt\xiadan.exe'
    } # 广发

    htong = {
        'account_no': '12345678',
        'password': '666666',
        'comm_pwd': '666666',
        'exe_path': r'D:\Utils\htong\xiadan.exe'
    } # 海通

    bdy = {
        'appId': '',
        'apiKey': '',
        'secretKey': ''
    } # 百度云 OCR https://cloud.baidu.com/product/ocr

    #t = Puppet(title='广发证券核新网上交易系统7.60')
    # 广发(标题为"广发证券核新网上交易系统7.65")拷贝数据会弹窗阻止。

    bot = Puppet().login(**gf)
    #bot.login(htong).wait(2).exit().login(htai).balance # 先登录海通，等2秒钟，退出，马上登录华泰查看余额
    #bot.cancel_order('000001', 'cancel')  # 取代cancel()方法。
