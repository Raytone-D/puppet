# -*- coding: utf-8 -*-
"木偶测试文件，请直接在命令行执行 python puppet\test.py"

if __name__ == '__main__':

    import platform

    from puppet.puppet import Puppet

    print('\n{}\nPython Version: {}'.format(platform.platform(), platform.python_version()))
    print('默认使用百度云OCR进行验证码识别')
    bdy = {
        'appId': '',
        'apiKey': '',
        'secretKey': ''
    } # 百度云 OCR https://cloud.baidu.com/product/ocr

    gf = {
        'account_no': '198800',
        'password': '',
        'comm_pwd': '',
        'client_path': r'D:\Utils\gfwt\xiadan.exe'
    } # 广发, 注意客户端路径不能错。

    htong = {
        'account_no': '12345678',
        'password': '666666',
        'comm_pwd': '666666',
        'client_path': r'D:\Utils\htong\xiadan.exe'
    } # 海通，注意客户端路径不能错。

    title='广发证券核新网上交易系统7.65' # 这个版本拷贝数据会弹窗询问
    bot = Puppet().login(**gf).wait(2)
    print(
        vars(bot), '\n',
        '余额:%s\n' % bot.balance,
        '持仓市值:%s\n' % bot.market_value)

    bot.buy('000001', '9.33', '100')
    bot.wait(2).cancel('000001')
    print('退出客户端...')
    bot.exit()
