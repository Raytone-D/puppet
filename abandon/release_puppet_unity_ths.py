__author__ =  '睿瞳深邃(https://github.com/Raytone-D)'
__project__ = "扯线木偶(puppet for THS trader)"
#增加账户id_btn = 1691
# coding: utf-8

import ctypes
from ctypes.wintypes import BOOL, HWND, LPARAM
from time import sleep
import win32clipboard as cp

WM_COMMAND, WM_SETTEXT, WM_GETTEXT, WM_KEYDOWN, WM_KEYUP, VK_CONTROL = \
    273,        12,         13,        256,        257,        17   # 消息命令
F1,  F2,  F3,  F4,  F5,  F6 = \
112, 113, 114, 115, 116, 117    # keyCode
op = ctypes.windll.user32
buffer = ctypes.create_unicode_buffer

def keystroke(hCtrl, keyCode, param=0):   # 击键
    op.PostMessageW(hCtrl, WM_KEYDOWN, keyCode, param)
    op.PostMessageW(hCtrl, WM_KEYUP, keyCode, param)

def get_data():
    sleep(0.3)    # 秒数关系到是否能复制成功。
    op.keybd_event(17, 0, 0, 0)
    op.keybd_event(67, 0, 0, 0)
    sleep(0.1)    # 没有这个就复制失败
    op.keybd_event(67, 0, 2, 0)
    op.keybd_event(17, 0, 2, 0)
    
    cp.OpenClipboard(None)
    raw = cp.GetClipboardData(13)
    data = raw.split()
    cp.CloseClipboard()
    return data
    
class unity():
    ''' 大一统协同交易 '''
    
    def __init__(self, hwnd):
        keystroke(hwnd, F6)    # 切换到双向委托
        self.buff = buffer(32)
        #            代码，价格，数量，买入，代码，价格，数量，卖出，全撤， 撤买， 撤卖
        id_members = 1032, 1033, 1034, 1006, 1035, 1058, 1039, 1008, 30001, 30002, 30003, \
                     32790, 1038, 1047, 2053, 30022   # 刷新，余额、表格、最后一笔、撤相同
        self.two_way = hwnd
        sleep(0.1)    # 按CPU的性能调整秒数(0.01~~0.5)，才能获取正确的self.two_way。
        for i in (59648, 59649):
            self.two_way = op.GetDlgItem(self.two_way, i)
        self.members = {i: op.GetDlgItem(self.two_way, i) for i in id_members}
        
    def buy(self, symbol, price, qty):   # 买入(B)
        op.SendMessageW(self.members[1032], WM_SETTEXT, 0, symbol)
        op.SendMessageW(self.members[1033], WM_SETTEXT, 0, price)
        op.SendMessageW(self.members[1034], WM_SETTEXT, 0, qty)
        op.PostMessageW(self.two_way, WM_COMMAND, 1006, self.members[1006])
    
    def sell(self, *args):    # 卖出(S)
        op.SendMessageW(self.members[1035], WM_SETTEXT, 0, symbol)
        op.SendMessageW(self.members[1058], WM_SETTEXT, 0, price)
        op.SendMessageW(self.members[1039], WM_SETTEXT, 0, qty)
        op.PostMessageW(self.two_way, WM_COMMAND, 1008, self.members[1008])

    def refresh(self):    # 刷新(F5)
        op.PostMessageW(self.two_way, WM_COMMAND, 32790, self.members[32790])
        
    def cancel(self, way=0):    # 撤销下单
        pass
        
    def cancelAll(self):    # 全撤(Z)
        op.PostMessageW(self.two_way, WM_COMMAND, 30001, self.members[30001])
        
    def cancelBuy(self):    # 撤买(X)
        op.PostMessageW(self.two_way, WM_COMMAND, 30002, self.members[30002])
        
    def cancelSell(self):    # 撤卖(C)
        op.PostMessageW(self.two_way, WM_COMMAND, 30003, self.members[30003])
        
    def cancelLast(self):    # 撤最后一笔，仅限华泰定制版有效
        op.PostMessageW(self.two_way, WM_COMMAND, 2053, self.members[2053])
        
    def cancelSame(self):    # 撤相同代码，仅限华泰定制版
        #op.PostMessageW(self.two_way, WM_COMMAND, 30022, self.members[30022])
        pass
        
    def balance(self):    # 可用余额
        op.SendMessageW(self.members[1038], WM_GETTEXT, 32, self.buff)
        return self.buff.value
        
    def position(self):    # 持仓(W)
        keystroke(self.two_way, 87)
        op.SetForegroundWindow(self.members[1047])
        return get_data()
    
    def tradeRecord(self):    # 成交(E)
        keystroke(self.two_way, 69)
        op.SetForegroundWindow(self.members[1047])
        return get_data()
        
    def orderRecord(self):    # 委托(R)
        keystroke(self.two_way, 82)
        op.SetForegroundWindow(self.members[1047])
        return get_data()
        
        
def finder(register):
    ''' 枚举所有可用的broker交易端并实例化 '''
    team = set()
    buff = buffer(32)
    @ctypes.WINFUNCTYPE(BOOL, HWND, LPARAM)
    def check(hwnd, extra):
        if op.IsWindowVisible(hwnd):
            op.GetWindowTextW(hwnd, buff, 32)
            if '交易系统' in buff.value:
                team.add(hwnd)
        return 1
    op.EnumWindows(check, 0)
    
    def get_nickname(hwnd):
        account = hwnd
        for i in 59392, 0, 1711:
            account = op.GetDlgItem(account, i)
        op.SendMessageW(account, WM_GETTEXT, 32, buff)
        return register.get(buff.value[-3:])
        
    return {get_nickname(hwnd): unity(hwnd) for hwnd in team if hwnd}
    
    
if __name__ == '__main__':

    myRegister = {'888': '股神','509': 'gf', '966': '女神', '167': '虚拟盘', '743': '西门吹雪'}   
    # 用来登录的号码（一般是券商客户号）最后3位数，不能有重复，nickname不能有重名！
    trader = finder(myRegister)
    if not trader:
        print("没发现可用的交易端。")
    else:
        #print(trader.keys())
        x = {nickname: broker.balance() for (nickname, broker) in trader.items()}
        print("可用余额：%s" %x)
        buy = '000078', '6.6', '300'
        #trader['虚拟盘'].buy(*buy)
        #p = trader['虚拟盘'].orderRecord()
        #p = trader['虚拟盘'].tradeRecord()
        p = trader['虚拟盘'].position()
        print(p)
        #trader['西门吹雪'].cancelLast()
    
