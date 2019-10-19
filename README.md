# 『扯线木偶』Puppet TraderAPI
## 突破交易桎梏！
&nbsp;
#### 需配合量化交易框架(hikyuu, quantaxis, rqalpha, vnpy, etc)使用。<br/>
#### 墙裂推荐使用实盘易，一站式量化交易解决方案 http://www.iguuu.com/e?x=19829 <br/>
<br/>
**快速入门**

```python
import puppet

# 自动登录账户, comm_pwd 是可选参数
accinfo = {
    'account_no': '你的证券账号',
    'password': '登录密码',
    'client_path': 'path/to/xiadan.exe',
    # 'comm_pwd': '通讯密码'
}

acc = puppet.login(accinfo)

# 绑定已登录账户
title = '广发证券核新网上交易系统7.65'
acc = puppet.Client(title=title)

# 交易
acc.buy('000001', 12.68, 100)
acc.sell('000001', 12.68, 100)
acc.cancel(choice='cancel_buy')
acc.cancel_sell()

# 查询
acc.position  # 持仓列表
acc.free_bal  # 可用余额
acc.assets  # 总资产
```
**使用环境**

1、核新同花顺股票交易客户端PC版。

2、Python3.5及以上，强烈推荐使用Anaconda3的最新版本。

3、(不推荐！)Linux平台需安装最新的Wine，环境设为WIN7，并安装Windows平台的Anaconda3。

**安装**

打开命令提示符或Windows PowerShell，然后执行：
```shell
pip install https://github.com/Raytone-D/puppet/archive/master.zip
```
或者
```shell
git clone https://github.com/Raytone-D/puppet.git
pip install -e puppet
```
**未实现的功能：**

融资融券、逆回购、基金盘后业务  
<br/>
**已实现的功能：**

make_heartbeat()  
随机制造心跳

login()  
账户登入客户端(支持通讯密码或验证码)

exit()  
关闭客户端

trade()          双向委托，【限价】或【市价】委托。

buy()           【限价】或【市价】买入。

sell()          【限价】或【市价】卖出。

cancel()        撤单

cancel_order()  撤单

cancel_all()    全撤

cancel_buy()    撤买

cancel_sell()   撤卖

raffle()        新股申购

account         登录账号

balance         可用余额

assets          总资产

position        持仓列表

market_value    持仓市值

cancelable      可撤委托

entrustment     当日委托

deals           当日成交

new             当日新股

bingo           中签查询，部分券商可查。

**技术说明：**

1、本项目使用User32.dll, Kernel32.dll所涵盖的win32 API。

2、按MSDN的API说明，win32 API支持WIN2000及以上版本，建议Win 7+。

**鸣谢：**

* https://github.com/hardywu/ 写了个rqalpha的接入模板PR，多谢支持！

///////////////////////做事有底线///////////////////////////////////////
