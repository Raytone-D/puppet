# -*- coding: utf-8 -*-
"木偶测试文件，请直接在命令行执行 python puppet\test.py"

if __name__ == '__main__':

    import platform
    import time

    import puppet

    print('\n{}\nPython Version: {}'.format(
        platform.platform(), platform.python_version()))
    print('默认使用百度云OCR进行验证码识别')
    print("\n注意！必须将client_path的值修改为你自己的交易客户端路径！\n")
    time.sleep(3)

    bdy = {
        'appId': '',
        'apiKey': '',
        'secretKey': ''
    } # 百度云 OCR https://cloud.baidu.com/product/ocr

    acc1 = {
        'account_no': '198800',
        'password': '123456',
        'comm_pwd': True,  # 模拟交易端必须为True
        'client_path': r'你的交易客户端目录\xiadan.exe'
    }

    raytone = {
        'account_no': '12345678',
        'password': '666666',
        #'comm_pwd': '666666',  # 没有通讯密码可以不写
        'client_path': r'D:\Utils\htong\xiadan.exe'
    }

    # 绑定已经登录的交易客户端，广发证券客户端需要额外指定标题
    title = '广发证券核新网上交易系统7.65'
    # acc = puppet.Client(title=title)

    # 自动登录交易客户端
    acc = puppet.login(acc1)
    # 如果取持仓数据有验证码弹窗，设置 acc.copy_protection = True

    print(
        vars(acc), '\n',
        '余额:%s\n' % acc.balance,
        '持仓市值:%s\n' % acc.market_value)

    acc.buy('510500', 4.688, 100)
    acc.wait(2).cancel('510500')
