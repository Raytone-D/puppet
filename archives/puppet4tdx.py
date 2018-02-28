"""
# the wrapper of A-shares local tdx client
"""
__project__ = 'Puppet'
__author__ = "睿瞳深邃(https://github.com/Raytone-D"
__version__ = "Fools' Day"

import ctypes
from functools import reduce
import time
import pywinauto

WM_SETTEXT = 12
WM_GETTEXT = 13
WM_SETCHECK = 241
WM_CLICK = 245
WM_KEYDOWN = 256
WM_KEYUP = 257
WM_COMMAND = 273

class Tdx:
    """ 通达信常量定义 """
    # 节点索引：0: '委托', 1: '撤单', 2: '资金股份', 3: '当日成交', 4: '新股申购', 5: '中签查询', 6: '批量申购',
    INIT = 'updown'
    GRID = 'SysListView32'
    CLS = 'TdxW_MainFrame_Class'
    TAG = ['对买对卖', '双向委托', '各种交易',
           '撤单', '撤单[F3]',
           '查询', '资金股份',
           '当日成交',
           '新股申购',
           '中签查询']    # 节点标签
    SUBTAG = ['资金股份', '资金股份F4',
              '当日成交', '当日成交查询',
              '新股申购',
              '中签查询', '新股中签缴款',
              '新股批量申购']    # 子节点标签。
    CUSTOM = {'招商定制': 'msctls_updown32',
              '定制1': '',
              'off': 1157}    # 交易面板定位。
    BUY = {'代码': 12005,
           '价格': 12006,
           '全部': 1495,
           '数量': 12007,
           '下单': 2010}
    SELL = {'代码': 2025,
            '价格': 12039,
            '全部': 3075,
            '数量': 3030,
            '下单': 3032}
    CANCEL = {'撤单': 1136,
              '全选': 14}
    NEW = {'全部': 2203,
           '申购': 11786,
           '弹窗': '新股申购确认',
           '确认': 7015}    # '新股代码': 12023,'申购价格': 12024,'最大可申': 2202,'申购数量': 12025,
    BATCH = {'申购': 39004,
             '弹窗': '新股组合申购确认',
             '确认': 7015}    # 批量申购
    BINGO = {'查询': 1140}
    PATH = (59648, 0, 0, 59648, 59649, 0)

api = ctypes.windll.user32
buff = ctypes.create_unicode_buffer(96)

def confirm_popup(idButton=7015, title='提示'):
    """ 确认弹窗 """
    time.sleep(0.5)
    popup = api.FindWindowW(0, title) or api.FindWindowW(0, '提示')
    if api.IsWindowVisible(popup):
        api.PostMessageW(popup, WM_COMMAND, idButton,
                         api.GetDlgItem(popup, idButton))
        return True
    else:
        print('没找到弹窗orz')
        return False

class Puppet:
    """
    # method: '委买': buy(), '委卖': sell(), '撤单': cancel(), '打新': raffle(), '下单': order(),
    # property: '帐号': account, '持仓': position, '可用余额': balance, '成交': deals,
    #           '可撤委托': cancelable, '新股': new, '中签': bingo
    """
    def __init__(self, main=None, clsName='TdxW_MainFrame_Class'):
        self.main = pywinauto.findwindows.find_window(class_name=clsName)
        app = pywinauto.Application().connect(handle=self.main)
        self._client = app.window(handle=self.main)
        self.tv = self._client['SysTreeView32']
        tag = [x.text() for x in self.tv.roots()]    # 节点标签
        self.tag = [x for x in Tdx.TAG if x in tag]    # 筛选
        self.tv.item(r'\对买对卖').click_input()
        self._trade = api.GetParent(self._client[Tdx.INIT])
        self.account = "暂不可用:("

    def _get_data(self, on=1):
        ''' 通达信SysListView32 '''
        lv = self._client['SysListView32']
        time.sleep(0.5)
        if on:
            raw = [x.text() for x in lv.items()]
            return list(zip(*[iter(raw)] * lv.column_count()))
        api.GetDlgItemTextW(api.GetParent(lv.handle), 1576, buff, 96)
        return dict([x.split(':') for x in buff.value.strip().split('  ')])

    def order(self, symbol, price, qty=0, way='买入'):
        """ 通达信下单 """
        self.tv.item(r'\对买对卖').click_input()
        _parts = Tdx.BUY if way == '买入' else Tdx.SELL
        if qty == 0:   # 暂时不可用。
            print("全仓{0}".format(way))
        print('{0:>>8} {1}, {2}, {3}'.format(way, symbol, price, qty))
        print('限价委托') if price else print('市价委托')
        time.sleep(0.5)    # 必须滴。
        api.SendMessageW(api.GetDlgItem(self._trade, _parts['代码']), WM_SETTEXT, 0, str(symbol))
        api.SendMessageW(api.GetDlgItem(self._trade, _parts['价格']), WM_SETTEXT, 0, str(price))
        api.SendMessageW(api.GetDlgItem(self._trade, _parts['数量']), WM_SETTEXT, 0, str(qty))
        time.sleep(0.5)
        api.PostMessageW(self._trade, WM_COMMAND, _parts['下单'],
                         api.GetDlgItem(self._trade, _parts['下单']))
        while True:
            if not confirm_popup(title='交易确认'):
                print('下单完成:)')
                break

    def buy(self, symbol, price, qty=0):
        return self.order(symbol, price, qty)

    def sell(self, symbol, price=0, qty=0):
        return self.order(symbol, price, qty, way='卖出')

    def cancel(self, symbol=None, number=None, way='卖出', comfirm=True):
        ''' 通达信撤单 '''
        self.tv.item(r'\撤单').click_input()
        print("撤单：{0}".format('>'*8))
        time.sleep(0.5)
        lv = self._client[Tdx.GRID]
        cancel = api.GetParent(lv.handle)
        if comfirm:
            temp = [{x[3]: [x[1], x[8]]} for x in self._get_data()]
            temp = [{x[way][0]: x[way][1]} for x in temp if x.get(way)]
            wanted = [temp.index(x) for x in temp if str(symbol) in x or str(number) in x.values()]
            if wanted:
                for x in wanted:
                    print('撤掉：{0} {1}'.format(way, temp[x]))
                    lv.item(x).click()
                    api.PostMessageW(cancel, WM_COMMAND, Tdx.CANCEL['撤单'],
                                     api.GetDlgItem(cancel, Tdx.CANCEL['撤单']))
                while True:
                    time.sleep(0.5)
                    if not confirm_popup():
                        print('撤单完成:)')
                        break
        return self._get_data()

    @property
    def cancelable(self):
        print('可撤委托: {0}'.format('$'*68))
        return self.cancel(way=False)

    @property
    def position(self):
        print('实时持仓: {0}'.format('$'*68))
        self.tv.item(r'\查询\资金股份').click_input()
        #time.sleep(0.5)    # 不一定需要用，备用。
        return self._get_data()

    @property
    def balance(self):
        print('资金明细: {0}'.format('$'*68))
        self.tv.item(r'\查询\资金股份').click_input()
        return self._get_data(on=False)

    @property
    def deals(self):
        print('当日成交: {0}'.format('$'*68))
        self.tv.item(r'\查询\当日成交').click_input()
        #time.sleep(0.5)    # 不一定需要用，备用。
        return self._get_data()

    @property
    def new(self):
        print('新股名单: {0}'.format('$'*68))
        self.tv.item(r'\新股申购\新股申购').click_input()
        return self._get_data()

    @property
    def bingo(self):
        print('新股中签: {0}'.format('$'*68))
        self.tv.item(r'\新股申购\中签查询').click_input()
        time.sleep(0.5)
        bingo = api.GetParent(self._client['SysListView32'].handle)
        api.PostMessageW(bingo, WM_COMMAND, Tdx.BINGO['查询'],
                         api.GetDlgItem(bingo, Tdx.BINGO['查询']))
        return self._get_data()

    def _batch(self):
        """ 新股批量申购，自动切换无需指定，只有华泰、银河、招商、广发等券商可用。 """
        self.tv.item(r'\新股申购\新股批量申购').click_input()
        print("新股批量申购：{0}".format('>'*8))
        print(self._get_data())
        time.sleep(0.5)
        batch = api.GetParent(self._client[Tdx.GRID].handle)
        api.PostMessageW(batch, WM_COMMAND, Tdx.BATCH['申购'],
                         api.GetDlgItem(batch, Tdx.BATCH['申购']))
        while True:
            time.sleep(0.5)
            if not confirm_popup(title=Tdx.BATCH['弹窗']):
                print('申购完毕！请查询可撤委托单:)')
                break

    def raffle(self, skip='', way=True):
        '通达信打新'
        tag = [x.text() for x in self.tv.item(r'\新股申购').children()]    # 子节点标签
        if '新股批量申购' in tag:
            self._batch()
        else:
            self.tv.item(r'\新股申购\新股申购').click_input()
            print("新股申购: {0}".format('>'*8))
            time.sleep(0.5)
            lv = self._client['SysListView32']
            raffle = api.GetParent(lv.handle)
            count = lv.item_count()
            _parts = [api.GetDlgItem(raffle, x) for x in Tdx.NEW.values()]
            for x in range(count):
                lv.item(x).click_input(double=True)
                api.PostMessageW(raffle, WM_COMMAND, Tdx.NEW['全部'], _parts[0])
                api.PostMessageW(raffle, WM_COMMAND, Tdx.NEW['申购'], _parts[1])
                while True:
                    time.sleep(0.5)
                    if not confirm_popup(title=Tdx.NEW['弹窗']):
                        print('申购完毕！请查询可撤委托单:)')
                        break

if __name__ == '__main__':

    trader = Puppet()
    print(trader.position)
    print(trader.balance)
    print(trader.cancelable)
    print(trader.deals)
    print(trader.new)
    print(trader.bingo)
    #trader.raffle()
    #trader.sell('002097', 11, 100)    # 00股票代码需要用字符串。
    trader.cancel(number=58)    # 默认撤卖单。symbol股票代码，number委托编号。
    
