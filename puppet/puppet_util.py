"""给木偶写的一些工具函数
"""
import configparser
import ctypes
import datetime
import os
import time

from ctypes.wintypes import WORD, DWORD


OPTIONS = {
    'GUI_CHEDAN_CONFIRM': 'no',
    'GUI_LOCK_TIMEOUT': '9999999999999999999999999999999999999999999999999999',
    'GUI_ORDER_CONFIRM': 'no',
    'GUI_REFRESH_TIME': '2',
    'GUI_WT_YDTS': 'yes',
    'SET_MCJG': 'empty',
    'SET_MCSL': 'empty',
    'SET_MRJG': 'empty',
    'SET_MRSL': 'empty',
    'SET_NOTIFY_DELAY': '1',
    'SET_POPUP_CJHB': 'yes',
    # 'SET_TOP_MOST': 'yes',
    # 'SYS_ZCSX_ENABLE': 'yes' #    自动刷新资产数据 修改无效，会强制恢复为no！
}

user32 = ctypes.windll.user32


def check_input_mode(h_edit, text='000001'):
    """获取 输入模式"""
    user32.SendMessageW(h_edit, 12, 0, text)
    time.sleep(0.3)
    return 'WM' if user32.SendMessageW(h_edit, 14, 0, 0) == len(text) else 'KB'


def get_root(key: list =['网上股票交易系统', '通达信']) -> tuple:
    from ctypes.wintypes import BOOL, HWND, LPARAM

    @ctypes.WINFUNCTYPE(BOOL, HWND, LPARAM)
    def callback(hwnd, lparam):
        user32.GetWindowTextW(hwnd, buf, 64)
        for s in key:
            if s in buf.value:
                handle.value = hwnd
                return False
        return True

    buf = ctypes.create_unicode_buffer(64)
    handle = ctypes.c_ulong()
    user32.EnumWindows(callback)
    return handle.value, buf.value


def get_today():
    return datetime.date.today()


def check_config(folder=None, encoding='gb18030'):
    """检查客户端xiadan.ini文件是否符合木偶的要求
    """
    with open(''.join([folder or os.getcwd(), r'\xiadan.ini']), encoding=encoding) as f:
        string = ''.join(('[puppet]\n', f.read()))

    conf = configparser.ConfigParser()
    conf.read_string(string)
    section = conf['SYSTEM_SET']

    print('推荐修改下列选项：')
    for key, value in OPTIONS.items():
        name, val = section.get(key).split(';')[3:5]
        if val != value:
            print(name, val, '改为', value, '\n')


if __name__ == "__main__":
    print('请在客户端目录内运行命令行！')
    check_config()
