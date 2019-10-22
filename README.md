# 本项目为学习WIN32 API而写，推荐一站式解决方案 http://www.iguuu.com/e?x=19829 <br/>
<br/>

**快速入门：**

```python
import puppet

# 自动登录账户, comm_pwd 是可选参数
accinfo = {
    'account_no': '你的账号',
    'password': '登录密码',
    'client_path': 'path/to/xiadan.exe',
    # 'comm_pwd': '通讯密码'
}

acc = puppet.login(accinfo)

# 绑定已登录账户
title = '广发证券核新网上交易系统7.65'
acc = puppet.Client(title=title)

# trade
acc.buy('000001', 12.68, 100)
acc.sell('000001', 12.68, 100)
acc.cancel(choice='cancel_buy')
acc.cancel_sell()

# query
acc.position  # 持仓列表
acc.free_bal  # 可用余额
acc.assets  # 总资产
```

**使用环境：**
1、Python3.5及以上，强烈推荐使用Anaconda3的最新版本。

2、(不推荐！)Linux平台需安装最新的Wine，环境设为WIN7，并安装Windows平台的Anaconda3。

**安装：**

打开命令提示符或Windows PowerShell，然后执行：

```shell
pip install https://github.com/Raytone-D/puppet/archive/master.zip
```

或者

```shell
git clone https://github.com/Raytone-D/puppet.git
pip install -e puppet
```

**技术说明：**

1、本项目使用User32.dll, Kernel32.dll所涵盖的win32 API。

2、按MSDN的API说明，win32 API支持WIN2000及以上版本，建议Win 7+。

**鸣谢：**

* https://github.com/hardywu/ 写了个rqalpha的接入模板PR，多谢支持！

///////////////////////做事有底线///////////////////////////////////////
