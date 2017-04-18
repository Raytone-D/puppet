"""
# 多账户打新专用脚本，支持v4+版本
# myRegister暂时没用上。暂时只支持同花顺交易端
"""
# coding: utf-8
import time
import os
import subprocess
import ctypes

from puppet_v4 import Puppet, switch_combo
from autologon import autologon

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
    for i in range(10):
        api.EnumWindows(check, keyword)
        time.sleep(3)
        if team:
            break
    return {Puppet(main) for main in team}

if __name__ == '__main__':

    myRegister = {'券商登录号': '自定义名称',
                  '617145470': '东方不败',
                  '20941552121212': '西门吹雪'}    # 交易端的登录帐号及昵称。
    keyword = '网上股票交易'

    autologon(target='同花顺交易')
    traders = find(keyword)
    for x in traders:
        popup = api.GetLastActivePopup(x.main)
        if popup:
            api.PostMessageW(popup, 273, 2, api.GetDlgItem(popup, 2))
        if api.IsWindowVisible(x.combo):
            for i in range(x.count):
                switch_combo(i, 2322, x.combo)
                time.sleep(1)
                x.raffle()
