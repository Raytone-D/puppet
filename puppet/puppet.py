# -*- coding: utf-8 -*-
"""
扯线木偶界面自动化应用编程接口(Puppet UIAutomation API)
技术群：624585416
"""
__author__ = "睿瞳深邃(https://github.com/Raytone-D)"
__project__ = 'Puppet'
__version__ = "0.5.9"
__license__ = 'MIT'

import ctypes
import time
import io
import subprocess
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
        'buy': 161,
        'sell': 162,
        'cancel_order': 163,
        'trade': 512,
        'entrustment': 168,
        'raffle':554,
        'bingo': 1070
    }
    PAGE = 59648, 59649
    PATH = {
        'account': (59392, 0, 1711),
        'balance': (1038,),
        'table': (1047, 200, 1047),
        'buy': (1032, 1033, 1034, '买入[B]', 1036, 1018),
        'sell':(1032, 1033, 1034, '卖出[S]', 1036, 1038),
        'cancel_order': (3348, '查询代码', '撤单')
    }

    buf_length = 32

    def __init__(self, title='网上股票交易系统5.0', main=None, **kwargs):
        print('Puppet TraderApi, version {}\n'.format(__version__))

        self.start = time.time()
        self._buf = ctypes.create_unicode_buffer(32)
        self.root = main
        self.init(title)

    def init(self, title=None, retry=1):
        self.root = self.root or user32.FindWindowW(0, title)
        if self.visible():
            print('木偶："正在热身，请稍候..."\n')
            user32.ShowOwnedPopups(self.root, False)
            #self.close_popup(delay=0.1) # 关闭"自动升级提示"弹窗
            self._container = {name: self._get_item(name) for name in self.NODE.keys()}
            self.members = {k: user32.GetDlgItem(self._container['trade'], v) for k, v in TWO_WAY.items()}

            print('木偶："我准备好了"\n')
            print("cost: %s\n" % (time.time() - self.start))

    def run(self, client_path):
        assert 'xiadan' in client_path and subprocess.os.path.isfile(client_path), '客户端路径错误'
        print('{} 正在尝试运行客户端({})...'.format(time.strftime('%Y-%m-%d %H:%M:%S %a'), client_path))

        BUTTON = {
            '连当前站点': '海通',
            '确定(&Y)': '华泰(广发)'
        }
        self.pid = subprocess.Popen(client_path).pid
        pid = ctypes.c_ulong()
        dlg = None
        for i in range(9):
            self.wait() # important
            hLogin = user32.FindWindowExW(dlg, None, '#32770', '用户登录')
            tid = user32.GetWindowThreadProcessId(hLogin, ctypes.byref(pid))
            #print(pid, self.pid)
            if pid.value == self.pid and self.visible(hLogin):
                self.wait()
                for label, value in BUTTON.items():
                    hButton = user32.FindWindowExW(hLogin, None, 'Button', label)
                    if self.visible(hButton):
                        break
                break
            else:
                dlg = hLogin

        assert hLogin, '客户端没有运行或者找不到用户登录窗口'
        assert hButton, '提交按钮标识错误'

        self.root = user32.GetParent(hLogin)
        self._hLogin = hLogin
        user32.GetWindowTextW(self.root, self._buf, 32)
        self.title = self._buf.value
        self.path = client_path
        self._label = label
        self.broker = value
        #self.wait(1) # important

        print("cost:", time.time() - self.start)

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
            ret = True
            if self.visible(handle):
                user32.GetClassNameW(handle, self._buf, self.buf_length)
                class_name = self._buf.value
                if class_name == 'Edit':
                    try:
                        text = next(lparam)
                        self.fill(handle, text)
                    except:
                        ret = False
            return ret

        if not comm_pwd:
            comm_pwd = self.verify(self.grab(), ocr)
        lparam = [account_no, password, comm_pwd]
        assert all(lparam), '用户登录参数不全'
        lparam = iter(lparam)
        user32.EnumChildWindows(self._hLogin, match, None)
        self.wait().commit(self._hLogin, self._label)

        for i in range(9):
            self.wait(1)
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
        hImage = user32.FindWindowExW(hParent or self._hLogin, None, 'Static', "")
        user32.GetWindowRect(hImage, ctypes.byref(rect))
        user32.SetForegroundWindow(hParent or self._hLogin)
        screenshot = ImageGrab.grab((rect.left, rect.top, rect.right*1.33, rect.bottom))
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
        return client.basicGeneral(image).get('words_result')[0]['words']

    def exit(self):
        "退出系统并关闭程序"
        assert self.visible(), "客户端没有登录"
        user32.PostMessageW(self.root, MSG['WM_CLOSE'], 0, 0)
        return self

    def switch(self, name=None): #root=None):
        node = {
            'account': 512,
            'balance': 512,
            'position': 512,
            'deals': 512,
            'new': 554,
            'cancelable': 163
        }.get(name) or self.NODE.get(name)
        print('page', name, node)
        #if isinstance(root, int):
        #    self.root = root
        assert self.visible(), "客户端已关闭或账户已登出"
        assert user32.SendMessageW(self.root, MSG['WM_COMMAND'], node, 0)
        return self

    def fill(self, editor, text):
        "fill in"
        user32.SendMessageW(editor, MSG['WM_SETTEXT'], 0, text)
        return self

    def commit(self, leader, label='确定'):
        committer = user32.FindWindowExW(leader, 0, 'Button', label)
        idCommitter = user32.GetDlgCtrlID(committer)
        user32.PostMessageW(leader, MSG['WM_COMMAND'], idCommitter, 0)
        return self

    def close_popup(self, button_label='以后再说', delay=0.5):
        for i in range(5):
            handle = user32.GetLastActivePopup(self.root)
            if handle:
                self.click_button(handle, button_label)
                return 1
            time.sleep(delay)

    def _get_item(self, name, sec=0.5):
        self.switch(name)
        time.sleep(sec)
        return reduce(user32.GetDlgItem, self.PAGE, self.root)

    def switch_tab(self, hCtrl, keyCode, param=0):   # 单击
        user32.PostMessageW(hCtrl, MSG['WM_KEYDOWN'], keyCode, param)
        time.sleep(0.1)
        user32.PostMessageW(hCtrl, MSG['WM_KEYUP'], keyCode, param)

    def query(self, category):
        self.switch(category)
        key = {
            'position': ord('W'),
            'deals': ord('E')
        }.get(category)
        if key:
            self.switch_tab(self._container['trade'], key)
        name = {
            'position': 'trade',
            'deals': 'trade',
            'cancelable': 'cancel_order',
            'new': 'raffle'
        }.get(category) or category
        handle = reduce(user32.GetDlgItem, self.PATH['table'], self._container[name])
        print(category, handle)
        return self.wait().copy_data(handle)

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
                    self.fill(hEdit, text).wait(0.1).commit(handle).wait(0.1)
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
        self._buf.value = ''  # False，待假成真
        for n in range(500):
            time.sleep(0.01)
            user32.SendDlgItemMessageW(container, id_item, MSG['WM_GETTEXT'], 64, self._buf)
            if self._buf.value:
                break

    def _order(self, container, id_items, *triple):
        self.fill_in(container, id_items[0], triple[0])  # 证券代码
        self._wait(container, id_items[-2])  # 证券名称
        self.fill_in(container, id_items[1], triple[1])  # 价格
        self._wait(container, id_items[-1])  # 可用数量
        self.fill_in(container, id_items[2], triple[2])  # 数量
        self.click_button(container, id_items[3])  # 下单按钮
        if len(str(triple[1]).split('.')[1]) == 3:  # 基金三位小数价格弹窗
            self.kill_popup(self.root)

    def buy(self, symbol, price, qty):
        self._order(self._container['buy'], self.NODE['buy'], symbol, price, qty)

    def sell(self, symbol, price, qty):
        self._order(self._container['sell'], self.NODE['sell'], symbol, price, qty)

    def buy2(self, symbol, price, qty, sec=0.3):   # 买入(B)
        user32.SendMessageW(self.members['买入代码'], MSG['WM_SETTEXT'], 0, str(symbol))
        time.sleep(0.1)
        user32.SendMessageW(self.members['买入价格'], MSG['WM_SETTEXT'], 0, str(price))
        time.sleep(0.1)
        user32.SendMessageW(self.members['买入数量'], MSG['WM_SETTEXT'], 0, str(qty))
        #user32.SendMessageW(self.members['买入'], MSG['BM_CLICK'], 0, 0)
        time.sleep(sec)
        user32.PostMessageW(self._container['trade'], MSG['WM_COMMAND'], TWO_WAY['买入'], 0)

    def sell2(self, symbol, price, qty, sec=0.3):    # 卖出(S)
        user32.SendMessageW(self.members['卖出代码'], MSG['WM_SETTEXT'], 0, str(symbol))
        time.sleep(0.1)
        user32.SendMessageW(self.members['卖出价格'], MSG['WM_SETTEXT'], 0, str(price))
        time.sleep(0.1)
        user32.SendMessageW(self.members['卖出数量'], MSG['WM_SETTEXT'], 0, str(qty))
        #user32.SendMessageW(self.members['卖出'], MSG['BM_CLICK'], 0, 0)
        time.sleep(sec)
        user32.PostMessageW(self._container['trade'], MSG['WM_COMMAND'], TWO_WAY['卖出'], 0)

    def refresh(self):    # 刷新(F5)
        user32.PostMessageW(self._container['trade'], MSG['WM_COMMAND'], TWO_WAY['刷新'], 0)

    def cancel_order(self, symbol=None, choice='cancel_all', symbolid=3348, nMarket=None, orderId=None):
        """撤销订单，choice选择操作的结果，默认“cancel_all”，可选“cancel_buy”、“cancel_sell”或"cancel"
            "cancel"是撤销指定股票symbol的全部委托。
        """
        hDlg = self._container['cancel_order']
        if symbol:
            self.fill_in(hDlg, symbolid, symbol)
            for i in range(10):
                time.sleep(0.3)
                self.click_button(hDlg, '查询代码')
                hButton = user32.FindWindowExW(hDlg, 0, 0, '撤单')
                # 撤单按钮的状态检查
                if user32.IsWindowEnabled(hButton):
                    break
        cases = {
            'cancel_all': '全撤(Z /)',
            'cancel_buy': '撤买(X)',
            'cancel_sell': '撤卖(C)',
            'cancel': '撤单'
        }
        self.click_button(hDlg, cases.get(choice))

    @property
    def account(self):
        handle = reduce(user32.GetDlgItem, self.PATH['account'], self.root)
        user32.SendMessageW(handle, MSG['WM_GETTEXT'], 32, self._buf)
        return self._buf.value

    @property
    def balance(self):
        self.switch('balance')
        user32.SendMessageW(self.members['可用余额'], MSG['WM_GETTEXT'], 32, self._buf)
        return float(self._buf.value)

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
        try:
            assert self.broker == '海通', '中签查询不支持{}证券！'.format(self.broker)
        except Exception as e:
            print(e, "\n该券商的中签查询未经测试，注意复查。")

        return self.query('bingo')

    def cancel_all(self):  # 全撤
        self.cancel_order()

    def cancel_buy(self):  # 撤买
        self.cancel_order(choice='cancel_buy')

    def cancel_sell(self):  # 撤卖
        self.cancel_order(choice='cancel_sell')

    def raffle(self, skip=False):    # 打新
        #user32.SendMessageW(self.root, MSG['WM_COMMAND'], NODE['新股申购'], 0)
        #self._raffle = reduce(user32.GetDlgItem, NODE['FORM'], self.root)
        #close_pop()    # 弹窗无需关闭，不影响交易。
        #schedule = self.query(self._raffle)
        ret = self.new
        if not ret:
            print("是日无新!")
            return ret
        self._raffle = reduce(user32.GetDlgItem, self.PATH['table'], self._container['raffle'])
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
            user32.SendMessageW(self._raffle_parts['可申购数量'], MSG['WM_GETTEXT'], 32, self._buf)
            if not int(self._buf.value):
                print('跳过零数量新股：{}'.format(symbol))
                continue
            user32.SendMessageW(self._raffle_parts['申购数量'], MSG['WM_SETTEXT'], 0, self._buf.value)
            time.sleep(0.3)
            user32.PostMessageW(self._raffle, MSG['WM_COMMAND'], NEW['申购'], 0)

        #user32.SendMessageW(self.root, MSG['WM_COMMAND'], NODE['双向委托'], 0)    # 切换到交易操作台
        return [new for new in self.cancelable if '配售申购' in new['操作']]

    def switch_combo(self, index, idCombo, hCombo):
        user32.SendMessageW(hCombo, MSG['CB_SETCURSEL'], index, 0)
        user32.SendMessageW(user32.GetParent(hCombo), MSG['WM_COMMAND'], MSG['CBN_SELCHANGE']<<16|idCombo, hCombo)

    def click_button(self, dialog, label):
        handle = user32.FindWindowExW(dialog, 0, 0, label)
        id_btn = user32.GetDlgCtrlID(handle)
        user32.PostMessageW(dialog, MSG['WM_COMMAND'], id_btn, 0)

    def fill_in(self, container, _id_item, _str):
        user32.SendDlgItemMessageW(container, _id_item, MSG['WM_SETTEXT'], 0, _str)

    def kill_popup(self, hDlg, name='是(&Y)'):
        for x in range(100):
            time.sleep(0.01)
            popup = user32.GetLastActivePopup(hDlg)
            if popup != hDlg and user32.IsWindowVisible(popup):
                yes = user32.FindWindowExW(popup, 0, 0, name)
                idYes = user32.GetDlgCtrlID(yes)
                user32.PostMessageW(popup, MSG['WM_COMMAND'], idYes, 0)
                print('popup has killed.')
                break

    def visible(self, hwnd=None):
        return user32.IsWindowVisible(hwnd or self.root)


if __name__ == '__main__':
    import platform
    print('\n{}\nPython Version: {}'.format(platform.platform(), platform.python_version()))
    gf = {
        'account_no': '666622xxxxxx',
        'password': '123456',
        'comm_pwd': '',
        'client_path': r'D:\Utils\gfwt\xiadan.exe'
    } # 广发

    htong = {
        'account_no': '12345678',
        'password': '666666',
        'comm_pwd': '666666',
        'client_path': r'D:\Utils\htong\xiadan.exe'
    } # 海通

    bdy = {
        'appId': '',
        'apiKey': '',
        'secretKey': ''
    } # 百度云 OCR https://cloud.baidu.com/product/ocr

    #t = Puppet(title='广发证券核新网上交易系统7.60')
    # 广发(标题为"广发证券核新网上交易系统7.65")拷贝数据会弹窗阻止。

    t = Puppet()
    t.login(**gf)
    #t.login(htong).wait(2).exit().login(htai).balance # 先登录海通，等2秒钟，退出，马上登录华泰查看余额
    #t.cancel_order('000001', 'cancel')  # 取代cancel()方法。
