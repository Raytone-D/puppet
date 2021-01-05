'''
option pal
需要将锁屏时间设为最大值
'''
__author__ = "睿瞳深邃(https://github.com/Raytone-D)"
__version__ = "1.2"
__license__ = 'MIT'


import ctypes
import io
import time
import functools

import keyboard as kb

from . import util


user32 = ctypes.windll.user32


class Option:
    '''期权'''
    send_order = 3001
    cancel = 3002
    undone= 3002
    position = 3003
    summary = 3003
    deal = 3004
    refresh = 3005

    def __init__(self, keyword: str = '期权宝') -> None:
        self.keyword = keyword
        if isinstance(keyword, str):
            self.initial(keyword)

    def initial(self, keyword):
        self.root = util.find_one(keyword)
        util.go_to_top(self.root)
        kb.send('F12')  # 切换到交易界面

        def get_page_by_trait(h_dialog: int, name: str, classname: str = 'Static') -> int:
            '''利用子控件的特征获取窗口页面句柄'''
            return user32.GetParent(util.find_single_handle(h_dialog, name, classname))

        self.h_switcher = get_page_by_trait(self.root, '下单', 'Button')
        h_menu = get_page_by_trait(self.root, '功能菜单', 'Button')
        h_treeview = user32.FindWindowExW(h_menu, None, 'SysTreeView32', None)
        rect = util.get_rect(h_treeview)
        self.xpos = rect[2] - int((rect[2] - rect[0]) * 0.5)
        self.ypos = rect[1] + 90  # 拆分申报 +90; 组合策略持仓查询 +100

        # 初始化页面
        for node in ('cancel', 'position', 'deal', 'send_order', 'portfolio'):
            self.switch(node)

        time.sleep(0.5)  # compatible with low frequency

        h_panel = get_page_by_trait(self.root, '买入开仓', 'Button')
        self.h_edit = user32.FindWindowExW(h_panel, None, 'Edit', None)
        handles = [user32.FindWindowExW(h_panel, None, 'Button', x) for x in ('重置', '买入', '卖出', '开仓', '平仓')]
        self.h_button = dict(zip(['reset', 'buy', 'sell', 'open', 'close'], handles))

        page = [get_page_by_trait(self.root, x) for x in ('撤单', '资金持仓', '当日成交')]
        data = [user32.FindWindowExW(x, None, 'Button', '输出') for x in page]
        self.h_split = util.find_single_handle(self.root, '拆分', 'Button')
        page = user32.GetParent(self.h_split)
        data.append(user32.FindWindowExW(page, None, 'Button', '输出'))
        self.h_export = dict(zip(['undone', 'position', 'deal', 'portfolio'], data))
        self.lv = util.find_single_handle(page, 'List1', 'SysListView32')

        self.buy_close = functools.partial(self.buy_open, action='buy_close')
        self.sell_open = functools.partial(self.buy_open, action='sell_open')
        self.sell_close = functools.partial(self.buy_open, action='sell_close')

    def switch(self, node: str, delay: int = 0.5):
        if util.go_to_top(self.root):
            if node in ('split', 'portfolio'):
                self.switch('send_order').switch('position')
                user32.SetCursorPos(self.xpos, self.ypos)
                user32.mouse_event(util.Msg.MOUSEEVENTF_LEFTDOWN | util.Msg.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                time.sleep(delay)
            else:
                id_button = getattr(self, node)
                user32.PostMessageW(self.h_switcher, util.Msg.WM_COMMAND, id_button, 0)
                time.sleep(delay)
            return self

    def query(self, category: str = 'summary'):
        '''查询交易数据
        category: 'summary', 'position', 'undone', 'deal', 'portfolio'
        '''
        self.switch(category)
        key = 'position' if category == 'summary' else category
        user32.PostMessageW(self.h_export[key], util.Msg.BM_CLICK, 0, 0)
        return self.normalize(self.handle_query_popup().export_csv(category))

    def handle_query_popup(self, times: int = 9):
        '''处理查询弹窗'''
        self.path = False
        buf = ctypes.create_unicode_buffer(64)
        for _ in range(times):
            h_popup: int = user32.GetLastActivePopup(self.root)
            if h_popup != self.root and user32.IsWindowVisible(h_popup):
                user32.SendMessageW(h_popup, util.Msg.WM_GETTEXT, 64, buf)
                if buf.value == '数据输出':
                    h_checkbox: int = user32.FindWindowExW(h_popup, None, 'Button', '主动打开输出的文件')
                    if user32.SendMessageW(h_checkbox, util.Msg.BM_GETSTATE, 0, 0):
                        user32.PostMessageW(h_checkbox, util.Msg.BM_CLICK, 0, 0)  # 不勾选

                    h_edit: int = user32.FindWindowExW(h_popup, None, 'Edit', None)
                    user32.SendMessageW(h_edit, util.Msg.WM_GETTEXT, 64, buf)  # 获取文件的绝对路径
                    self.path: str = buf.value

                    kb.send('\r')
                    continue
                elif buf.value == '数据导出':
                    kb.send('\r')
                    break
                elif buf.value == '提示':
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
                        text = dict(summary=x, position=y.replace('没有查询到相应的数据！', '', 1))[category]
                    return text.expandtabs().strip()

    def normalize(self, text: str):
        '''将文本数据转为DataFrame格式'''
        return util.pd.read_csv(io.StringIO(text), sep='\s+') if text else util.pd.DataFrame()

    def buy_open(self, symbol, price=None, quantity=None, action='buy_open'):
        '''
        price: None 取默认值 卖1
        quantity: None 取默认值 1
        '''
        self.switch('send_order')

        user32.PostMessageW(self.h_button['reset'], util.Msg.BM_CLICK, 0, 0)  # 清除默认的证券代码
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
        user32.PostMessageW(self.h_button[act], util.Msg.BM_CLICK, 0, 0)
        user32.PostMessageW(self.h_button[open_or_close], util.Msg.BM_CLICK, 0, 0)  # 开平标志
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
        return self.handle_trading_popup()

    def handle_trading_popup(self, times: int = 9):
        '''处理下单弹窗'''
        rtn = False
        label = ('买入开仓', '买入平仓', '卖出开仓', '卖出平仓')
        buf = ctypes.create_unicode_buffer(64)
        for _ in range(times):
            h_popup: int = user32.GetLastActivePopup(self.root)
            if h_popup != self.root and user32.IsWindowVisible(h_popup):
                user32.SendMessageW(h_popup, util.Msg.WM_GETTEXT, 64, buf)
                if buf.value == '快速下单':
                    kb.send('\r')
                elif buf.value in label:
                    kb.send('\r')
                    break
                elif buf.value == '提示':
                    rtn = True
                    kb.send('\r')
                    break
                elif buf.value == '拆分申报':
                    h_edit = user32.FindWindowExW(h_popup, None, 'Edit', None)
                    user32.SendMessageW(h_edit, util.Msg.WM_SETTEXT, 0, self.qty)
                    kb.send('\r')
                elif buf.value == '警告':
                    rtn = True
                    kb.send('\r')
                    break

            time.sleep(0.1)
        kb.send('esc')
        return rtn

    def split_decl(self, strategy_id: int, quantity: int, delay=0.1):
        '''拆分申报'''
        df = self.query('portfolio')
        if not df.empty:
            rect = util.get_rect(self.lv)
            user32.SetCursorPos(rect[0] + 9, rect[1] + 29)
            user32.mouse_event(util.Msg.MOUSEEVENTF_LEFTDOWN | util.Msg.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            time.sleep(delay)

            for x in range(df['序号'].to_list().index(strategy_id)):
                if x > 0:
                    kb.send(kb.KEY_DOWN)

            self.qty = str(quantity)
            user32.PostMessageW(self.h_split, util.Msg.BM_CLICK, 0, 0)
            return self.handle_trading_popup()


if __name__ == "__main__":
    opt = Option()
    time.sleep(3)
    df = opt.query('position')
    print(df)
