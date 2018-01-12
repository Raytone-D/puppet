"""
    测试一个或者多个账号是否可用。
    如果用来做多账号交易，需要先确认账号，再下单。
"""
# coding: utf-8

__author__ = '睿瞳深邃'
__version__ = '0.2'

import ctypes

from puppet_v4 import Puppet

api = ctypes.windll.user32
buff = ctypes.create_unicode_buffer(32)


def find(keyword='交易系统'):
    """ 枚举所有已登录的交易端 """
    clients = set()
    
    @ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p, ctypes.c_wchar_p)
    def check(hwnd, keyword):
        """ callback function, 用标题栏关键词筛选 """
        if api.IsWindowVisible(hwnd) and api.GetWindowTextW(hwnd, buff, 32) > 6 and keyword in buff.value:
            clients.add(hwnd)
        return 1

    api.EnumWindows(check, keyword)

    return {Puppet(c) for c in clients} if clients else None


myRegister = {
    '617145470': '东方不败',
    '20941552121212': '西门吹雪'
}    # 客户端的登录账号及自定义代号。
keyword = '广发证券'  # 自定义关键词
traders = find()
if traders:
    for x in traders:
        print(x.account)
        print(x.position)
        #print(x.new)
        #x.raffle()
else:
    print('<木偶：没找到可用的客户端>')
