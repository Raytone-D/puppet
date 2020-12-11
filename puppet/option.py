'''
option pal
需要将锁屏时间设为最大值
'''
__author__ = "睿瞳深邃(https://github.com/Raytone-D)"
__version__ = "0.6"
__license__ = 'MIT'


import ctypes
import io
import time
import functools

import keyboard as kb

from puppet import util


user32 = ctypes.windll.user32


def go_to_top(h_root: int):
    for x in range(99):
        if user32.GetForegroundWindow() == h_root:
            return True
        user32.SwitchToThisWindow(h_root, True)
        time.sleep(0.01) # DON'T REMOVE!


def find_single_handle(h_dialog, keyword: str = '', classname='Static') -> int:
    from ctypes.wintypes import BOOL, HWND, LPARAM

    @ctypes.WINFUNCTYPE(BOOL, HWND, LPARAM)
    def callback(hwnd, lparam):
        user32.SendMessageW(hwnd, util.Msg.WM_GETTEXT, 64, buf)
        user32.GetClassNameW(hwnd, buf1, 64)
        if buf.value == keyword and buf1.value == classname:
            handle.value = hwnd
            return False
        return True

    buf = ctypes.create_unicode_buffer(64)
    buf1 = ctypes.create_unicode_buffer(64)
    handle = ctypes.c_ulong()
    user32.EnumChildWindows(h_dialog, callback)
    return handle.value


class Option:
    '''期权'''
    send_order = 3001
    cancel = 3002
    undone= 3002
    position = 3003
    summary = 3003
    deal = 3004
    refresh = 3005
    buy = 6006
    sell = 6007
    open = 6009
    close = 6010
    FOK = 5027
    备兑 = 5025

    def __init__(self, keyword: str = '期权宝') -> None:
        self.keyword = keyword
        if isinstance(keyword, str):
            self.initial(keyword)

    def initial(self, keyword):
        self.root = util.find_one(keyword)
        go_to_top(self.root)  # 置顶
        kb.send('F12')

        def get_page_by_trait(h_dialog: int, name: str, classname: str = 'Static') -> int:
            '''利用子控件的特征获取窗口页面句柄'''
            return user32.GetParent(find_single_handle(h_dialog, name, classname))

        self.h_switcher = get_page_by_trait(self.root, '下单', 'Button')

        # 初始化页面
        for node in ('cancel', 'position', 'deal', 'send_order'):
            self.switch(node)

        self.h_panel = get_page_by_trait(self.root, '买入开仓', 'Button')
        self.h_reset = user32.FindWindowExW(self.h_panel, None, 'Button', '重置')
        self.h_edit = user32.FindWindowExW(self.h_panel, None, 'Edit', None)

        page = [get_page_by_trait(self.root, x) for x in ('撤单', '资金持仓', '当日成交')]
        data = [user32.FindWindowExW(x, None, 'Button', '输出') for x in page]
        self.h_export = dict(zip(['undone', 'position', 'deal'], data))

        self.buy_close = functools.partial(self.buy_open, action='buy_close')
        self.sell_open = functools.partial(self.buy_open, action='sell_open')
        self.sell_close = functools.partial(self.buy_open, action='sell_close')

    def switch(self, node: str, delay: int = 0.5):
        if go_to_top(self.root):
            id_button = getattr(self, node)
            user32.PostMessageW(self.h_switcher, util.Msg.WM_COMMAND, id_button, 0)
            time.sleep(delay)

    def query(self, category: str = 'summary'):
        '''查询交易数据
        category: 'summary', 'position', 'undone', 'deal'
        '''
        self.switch(category)
        key = 'position' if category == 'summary' else category
        user32.PostMessageW(self.h_export[key], util.Msg.BM_CLICK, 0, 0)
        return self.normalize(self.handle_popup().export_csv(category))

    def handle_popup(self, times: int = 9):
        self.path = False
        buf = ctypes.create_unicode_buffer(64)
        for _ in range(times):
            h_popup = user32.GetLastActivePopup(self.root)
            user32.SendMessageW(h_popup, util.Msg.WM_GETTEXT, 64, buf)
            if buf.value == '数据输出' and user32.IsWindowVisible(h_popup):
                h_checkbox: int = user32.FindWindowExW(h_popup, None, 'Button', '主动打开输出的文件')
                if user32.SendMessageW(h_checkbox, util.Msg.BM_GETSTATE, 0, 0):
                    user32.PostMessageW(h_checkbox, util.Msg.BM_CLICK, 0, 0)  # 不勾选
                h_edit = user32.FindWindowExW(h_popup, None, 'Edit', None)
                user32.SendMessageW(h_edit, util.Msg.WM_GETTEXT, 64, buf)  # 获取文件的绝对路径
                self.path = buf.value
                kb.send('\r')
                continue
            elif buf.value == '数据导出' and user32.IsWindowVisible(h_popup):
                kb.send('\r')
                break
            elif buf.value == '提示' and user32.IsWindowVisible(h_popup):
                kb.send('\r')
                break
            time.sleep(0.1)
        return self

    def export_csv(self, category: str) -> str:
        if self.path:
            for _ in range(9):
                try:
                    with open(self.path) as f:
                        text = f.read()
                except Exception as e:
                    print(e)
                    time.sleep(0.1)
                else:
                    if category in ('summary', 'position'):
                        x, y = text.split('\n\n', 1)
                        text = dict(summary=x, position=y)[category]
                    return text.expandtabs().strip()

    def normalize(self, text: str):
        '''将文本数据转为DataFrame格式'''
        return util.pd.read_csv(io.StringIO(text), sep='\s+')\
            .drop(columns=['序号']).dropna(how='all') if text else util.pd.DataFrame()

    def buy_open(self, symbol, price=None, quantity=None, action='buy_open'):
        '''
        price: None 取默认值 卖1
        quantity: None 取默认值 1
        '''
        self.switch('send_order')

        user32.PostMessageW(self.h_reset, util.Msg.BM_CLICK, 0, 0)  # 清除默认的证券代码
        for _ in range(9):
            if user32.SendMessageW(self.h_edit, util.Msg.WM_GETTEXTLENGTH, 0, 0) == 0:
                break
            time.sleep(0.1)

        kb.write(symbol)
        for _ in range(9):
            if user32.SendMessageW(self.h_edit, util.Msg.WM_GETTEXTLENGTH, 0, 0) == len(symbol):
                break
            time.sleep(0.05)

        act, open_or_close = action.split('_')
        id_act = getattr(self, act)
        id_ooc = getattr(self, open_or_close)
        h_act = user32.GetDlgItem(self.h_panel, id_act)
        h_ooc = user32.GetDlgItem(self.h_panel, id_ooc)
        user32.SendMessageW(h_act, util.Msg.BM_CLICK, 0, 0)
        user32.PostMessageW(h_ooc, util.Msg.BM_CLICK, 0, 0)  # 开平标志
        if open_or_close == 'open':
            kb.send('DOWN')
        kb.send('DOWN')
        kb.send('DOWN')  # 跳过报价方式，默认限价
        if price != None:
            kb.write(str(price))
        kb.send('\r')
        if quantity != None:
            kb.write(str(quantity))
        kb.send('\r')
        return self._handle_popup(action)

    def _handle_popup(self, action, times: int = 9):
        rtn = False
        buf = ctypes.create_unicode_buffer(64)
        for _ in range(times):
            h_popup = user32.GetLastActivePopup(self.root)
            user32.SendMessageW(h_popup, util.Msg.WM_GETTEXT, 64, buf)
            if buf.value == '快速下单' and user32.IsWindowVisible(h_popup):
                kb.send('\r')
            elif buf.value in ('买入开仓', '买入平仓', '卖出开仓', '卖出平仓') and user32.IsWindowVisible(h_popup):
                kb.send('\r')
                break
            elif buf.value == '提示' and user32.IsWindowVisible(h_popup):
                rtn = True
                kb.send('\r')
                break
            time.sleep(0.1)
        kb.send('esc')
        return rtn

if __name__ == "__main__":
    opt = Option()
    time.sleep(3)
    df = opt.query('position')
    print(df)
