"""puppetrader for ths client uniform"""
__author__ =  '睿瞳深邃(https://github.com/Raytone-D)'
__project__ = '扯线木偶(puppetrader for ths client unity)'
__version__ = "0.3.5"
'推荐使用：anaconda3 最新版，或者Python >= 3.6'
# coding: utf-8

import ctypes
from functools import reduce
import time
import pyperclip
import json

WM_COMMAND, WM_SETTEXT, WM_GETTEXT, WM_KEYDOWN, WM_KEYUP = \
    273,        12,         13,        256,        257           # 命令
    
F1,   F2,   F3,    F4,    F5,    F6 = \
112,  113,  114,   115,   116,   117       # keyCode(按键代码)

op = ctypes.windll.user32
wait_a_second = lambda sec= 0.1: time.sleep(sec)

def keystroke(hCtrl, keyCode, param=0):   # 单击
    op.PostMessageW(hCtrl, WM_KEYDOWN, keyCode, param)
    op.PostMessageW(hCtrl, WM_KEYUP, keyCode, param)
    
class unity():
    ''' 多账户交易集中处理 '''
    
    def __init__(self, main):
    
        self.main = main
        keystroke(main, F6)    # 切换到双向委托
        wait_a_second()    # 可调整区间值(0.01~0.5)
        
        self.buff = ctypes.create_unicode_buffer(32)
        #            代码，价格，数量，买入，代码，价格，数量，卖出，全撤， 撤买， 撤卖
        id_members = 1032, 1033, 1034, 1006, 1035, 1058, 1039, 1008, 30001, 30002, 30003, \
                     32790, 1038, 1047, 2053, 30022, 1019   # 刷新，余额、表格、最后一笔、撤相同
        self.path_custom = 1047, 200, 1047
                     
        self.two_way = reduce(op.GetDlgItem, (59648, 59649), main)
        self.members = {i: op.GetDlgItem(self.two_way, i) for i in id_members}
        self.custom = reduce(op.GetDlgItem, self.path_custom, self.two_way)
        
        # 获取登录账号
        self.account = reduce(op.GetDlgItem, (59392, 0, 1711), main)
        op.SendMessageW(self.account, WM_GETTEXT, 32, self.buff)
        self.account = self.buff.value
        # 撤单工具条
        self.id_toolbar = {'全选': 1098, \
                           '撤单': 1099, \
                           '全撤': 30001, \
                           '撤买': 30002, \
                           '撤卖': 30003, \
                           '填单': 3348, \
                           '查单': 3349} #'撤尾单': 2053, '撤相同': 30022}    # 华泰独有
                      
        op.SendMessageW(main, WM_COMMAND, 163, 0)    # 切换到撤单操作台
        wait_a_second()
        self.cancel_panel = reduce(op.GetDlgItem, (59648, 59649), main)
        self.cancel_toolbar = {k: op.GetDlgItem(self.cancel_panel, v) for k, v in self.id_toolbar.items()}
        keystroke(main, F6)    # 切换到双向委托

    def buy(self, symbol, price, qty):   # 买入(B)
        # buy = order('b')
        op.SendMessageW(self.members[1032], WM_SETTEXT, 0, symbol)
        op.SendMessageW(self.members[1033], WM_SETTEXT, 0, price)
        op.SendMessageW(self.members[1034], WM_SETTEXT, 0, qty)
        op.PostMessageW(self.two_way, WM_COMMAND, 1006, self.members[1006])
    
    def sell(self, symbol, price, qty):    # 卖出(S)
        # buy = order('s')
        op.SendMessageW(self.members[1035], WM_SETTEXT, 0, symbol)
        op.SendMessageW(self.members[1058], WM_SETTEXT, 0, price)
        op.SendMessageW(self.members[1039], WM_SETTEXT, 0, qty)
        op.PostMessageW(self.two_way, WM_COMMAND, 1008, self.members[1008])

    def refresh(self):    # 刷新(F5)
        op.PostMessageW(self.two_way, WM_COMMAND, 32790, self.members[32790])
        
    def cancel_order(self, symbol=''):    # 撤单
        op.SendMessageW(main, WM_COMMAND, 163, 0)    # 切换到撤单操作台
        if symbol:
            op.SendMessageW(self.cancel_toolbar['填单'], WM_SETTEXT, 0, symbol)
            sleep(0.1)    # 必须有
            op.PostMessageW(self.cancel_panel, WM_COMMAND, self.id_toolbar['查单'], self.cancel_toolbar['查单'])
            op.PostMessageW(self.cancel_panel, WM_COMMAND, self.id_toolbar['撤单'], self.cancel_toolbar['撤单'])
            keystroke(self.main, F6)    # 必须返回双向委托操作台!
            
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
        
    def get_data(self, key='W'):
        "将CVirtualGridCtrl|Custom<n>的数据复制到剪贴板，默认取持仓记录"
        
        keystroke(self.two_way, ord(key))    # 切换到持仓('W')、成交('E')、委托('R')
        wait_a_second()    # 等待券商的数据返回...
        op.SendMessageW(self.custom, WM_COMMAND, 57634, self.path_custom[-1])    # background mode
        
        return pyperclip.paste()
        
def finder():
    """ 枚举所有已登录的交易端并将其实例化 """
    
    team = set()
    buff = ctypes.create_unicode_buffer(32)
    
    @ctypes.WINFUNCTYPE(ctypes.wintypes.BOOL, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
    def check(hwnd, extra):
        if op.IsWindowVisible(hwnd):
            op.GetWindowTextW(hwnd, buff, 32)
            if '交易系统' in buff.value:
                team.add(hwnd)
        return 1
    op.EnumWindows(check, 0)
    
    return {unity(main) for main in team if team}
    
def to_dict(raw):
    x = [row.split() for row in raw.splitlines()]
    data = {}
    for y in x[1:]:
        data.update({y[0]: dict(zip(x[0], y))})
    return data
    
if __name__ == '__main__':

    myRegister = {'券商登录号': '自定义名称',   \
                  '617145470': '东方不败', \
                  '20941552121212': '西门吹雪'}
    
    ret = finder()
    
    if ret:
        # 如果没取到余额，尝试修改_init_函数里面sleep的值，或者查余额的id是不是变了。
        trader = {myRegister[broker.account]: broker for broker in ret}    # 给账户一个易记的外号
        trader1 = {broker.account[-3:]: broker.balance() for broker in ret} # 以登录号3位尾数作代号
        profile = {solo: {"交易帐号": trader[solo].account, \
                          "可用余额": trader[solo].balance()} \
                   for solo in trader}
                   
        print(profile)
        #print(json.dumps(profile, indent=4, ensure_ascii=False, sort_keys=True))
        
        raw = trader['东方不败'].get_data()    # 只能“大写字母”，小写字母THS会崩溃，无语！
        print(raw)
        #print(to_dict(raw))
        
        raw = trader['东方不败'].get_data('R')
        print(raw)
        #print(json.dumps(to_dict(raw), indent=4, ensure_ascii=False))
        trader['东方不败'].cancel_order('')
    
    else: print("老板，没发现已登录的交易端！")
