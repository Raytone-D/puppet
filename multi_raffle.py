""" 多账户打新专用脚本 支持v4+版本 """
__author__ = '睿瞳深邃'
__version__ = '0.1'

# coding: utf-8
from puppet_v4 import Puppet
import ctypes
api = ctypes.windll.user32

buff = ctypes.create_unicode_buffer(32)
team = set()
def find(keyword):
    """ 枚举所有已登录的交易端 """
    @ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p, ctypes.c_wchar_p)
    def check(hwnd, keyword):
        """ 筛选 """
        if api.IsWindowVisible(hwnd)\
            and api.GetWindowTextW(hwnd, buff, 32) > 6 and keyword in buff.value:
            team.add(hwnd)
        return 1
    api.EnumWindows(check, keyword)
    return {Puppet(main) for main in team}

myRegister = {'券商登录号': '自定义名称',
              '617145470': '东方不败',
              '20941552121212': '西门吹雪'}    # 交易端的登录帐号及昵称。
keyword = '网上股票交易'

traders = find(keyword)
for x in traders:
    print(x.account)
    #print(x.new)
    x.raffle()

# myRegister暂时没用上。暂时只支持同花顺交易端。
