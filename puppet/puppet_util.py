"""给木偶写的一些工具函数
"""
import configparser
import ctypes
import datetime
import io
import os
import threading
import time

from ctypes.wintypes import WORD, DWORD

try:
    import pandas as pd
except Exception as e:
    print(e)

try:
    import keyboard
    from keyboard import write as fill
except Exception as e:
    print(e)
    from pywinauto import keyboard
    from pywinauto.keyboard import send_keys as fill


class Msg:

    WM_SETTEXT = 12
    WM_GETTEXT = 13
    WM_CLOSE = 16
    WM_KEYDOWN = 256
    WM_KEYUP = 257
    WM_COMMAND = 273
    BM_CLICK = 245
    CB_GETCOUNT = 326
    CB_SETCURSEL = 334
    CBN_SELCHANGE = 1


COLNAMES = {
    '证券代码': 'symbol',
    '证券名称': 'name',
    '股票余额': 'quantity',  # 海通证券
    '当前持仓': 'quantity',  # 银河证券
    '可用余额': 'leftover',
    '冻结数量': 'frozen',
    '盈亏': 'profit',
    '参考盈亏': 'profit',  # 银河证券
    '浮动盈亏': 'profit',  # 广发证券
    '市价': 'price',
    '市值': 'amount',
    '参考市值': 'amount',  # 银河证券
    '最新市值': 'amount',  # 国金|平安证券
    '成交时间': 'time',
    '成交日期': 'date',
    '成交数量': 'quantity',
    '成交均价': 'price',
    '成交价格': 'price',
    '成交金额': 'amount',
    '成交编号': 'id',
    '申报时间': 'time',
    '委托日期': 'date',
    '委托时间': 'time',
    '委托价格': 'order_price',
    '委托数量': 'order_qty',
    '合同编号': 'order_id',
    '委托编号': 'order_id',  # 银河证券
    '委托状态': 'status',
    '操作': 'op',
    '发生金额': 'total',
    '手续费': 'commission',
    '印花税': 'tax',
    '其他杂费': 'fees'
}

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


def capture_popup(time_interval=0.5):
    ''' 弹窗截获 '''

    def capture(time_interval):
        while True:
            # secs = random.uniform(remainder/2, remainder)
            # 若在休眠期间心跳印记没被修改，则刷新页面并修改心跳印记
            time.sleep(secs)

    threading.Thread(
        target=capture,
        kwargs={'time_interval': time_interval},
        name='capture_popup',
        daemon=True).start()


def normalize(string: str, to_dict=False):
    '''标准化输出交易数据'''
    df = pd.read_csv(io.StringIO(string), sep='\t', dtype={'证券代码': str})
    df.drop(columns=[x for x in df.columns if x not in COLNAMES], inplace=True)
    df.columns = [COLNAMES.get(x) for x in df.columns]
    if 'amount' in df.columns:
        df['ratio'] = (df['amount'] / df['amount'].sum()).round(2)
    return df.to_dict('list') if to_dict else df


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


def get_text(h_parent, text_id):
    '获取控件文本内容'
    buf = ctypes.create_unicode_buffer(64)
    user32.SendDlgItemMessageW(h_parent, text_id, Msg.WM_GETTEXT, 64, buf)
    return buf.value.rstrip('%')


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
