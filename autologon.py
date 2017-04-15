"""
# autologon.py
# 目前仅支持同花顺官方的独立交易端的“多帐号”登录模式。
"""
__author__ = '睿瞳深邃'
__version__ = '0.1'

# coding: utf-8
import os
import subprocess
import time
import ctypes

api = ctypes.windll.user32

def autologon(target=None):
    " 自动登录同花顺交易客户端 "
    buff = ctypes.create_unicode_buffer(32)
    path = os.path.split(os.path.realpath(__file__))[0]
    for lnk in os.listdir(path):
        if target in lnk:
            subprocess.Popen(os.path.join(path, lnk), shell=True)
    for i in range(10):
        main = api.FindWindowW(0, '网上股票交易系统5.0')
        if not main:
            time.sleep(1)
        else: break
    if not api.IsWindowVisible(main):
        popup = api.GetLastActivePopup(main)
        logon = api.GetDlgItem(popup, 1015)
        api.GetWindowTextW(logon, buff, 32)
        if buff.value != '一键登录':
            api.PostMessageW(popup, 273, 1010, api.GetDlgItem(popup, 1010))
        api.PostMessageW(popup, 273, 1015, logon)

if __name__ == '__main__':

    autologon('同花顺交易.lnk')
