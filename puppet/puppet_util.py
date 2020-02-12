"""给木偶写的一些工具函数
"""
import configparser
import datetime
import os


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
