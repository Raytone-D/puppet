# -*- coding: utf-8 -*-
"木偶测试文件，请直接在命令行执行 python puppet\test.py"

if __name__ == '__main__':

    import platform
    import time

    from puppet.puppet import Puppet

    print('\n{}\nPython Version: {}'.format(platform.platform(), platform.python_version()))
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

    title='广发证券核新网上交易系统7.65' # 这个版本拷贝数据会弹窗询问
    # quant = Puppet() # 已经登录的，标题是"网上股票交易系统5.0"的交易客户端。
    # quant = Puppet(title) # 广发需要指定标题
    quant = Puppet().login(**acc1).wait(2)

    print(
        vars(quant), '\n',
        '余额:%s\n' % quant.balance,
        '持仓市值:%s\n' % quant.market_value)

    quant.buy('510500', 4.688, 100)
    quant.wait(2).cancel('510500')
