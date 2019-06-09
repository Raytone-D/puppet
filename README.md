# 『扯线木偶』Puppet TraderAPI
### 突破交易桎梏！
&nbsp;
#### puppet扯线木偶目前仅适用于独立交易端核新版(即THS)，需配合量化交易框架(hikyuu, quantaxis, rqalpha, vnpy, etc)使用。<br/>
#### 墙裂推荐使用实盘易，一站式量化交易解决方案 http://www.iguuu.com/e?x=19829 <br/>
#### 实现了和股票交易客户端相同的【买卖撤查】功能。暂不支持【融资融券】交易功能。
**快速入门**
```python
from puppet import Client

# 自动登录账户
accinfo = {
    'account_no': '你的账号',
    'password': '登录密码',
    'client_path': 'path/to/xiadan.exe',
    'comm_pwd': '通讯密码'
}  # comm_pwd 是可选参数
acc = Client().login(**accinfo)

# 绑定已登录账号
title = '广发证券核新网上交易系统7.65'
acc = Client().bind(title).init()

# 交易
acc.buy('000001', 12.68, 100)
acc.sell('000001', 12.68, 100)
acc.cancel(choice='cancel_buy')
acc.cancel_sell()

# 查询
acc.position
acc.balance
acc.assets
```

**安装**

在命令提示符(Command Prompt)后面：
```shell
git clone https://github.com/Raytone-D/puppet.git

python /path/to/puppet/setup.py install
```
或者浏览 https://github.com/Raytone-D/puppet
点击页面的绿色按钮[Clone or download]，选择[Download ZIP]，
```shell
pip install /path/to/puppet-master.zip
```

**未实现的功能：**

逆回购、基金盘后业务  

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


**使用环境：**

1、同花顺 股票交易客户端，留意【部分最新】版本不能复制持仓或委托数据。推荐备用版本: https://pan.baidu.com/s/1radY1fI 密码: m9db

1、要求用Python3.4及以上，强烈推荐用于科学计算的Python发行版【Anaconda3】的【最新】版本。

2、Win 2k+，Windows下免安装、零配置

3、Linux需安装最新的Wine，环境设为WIN7，并安装Win平台的Anaconda3。

4、下载Anaconda3，推荐用清华镜像：https://mirrors.tuna.tsinghua.edu.cn/anaconda/archive/


**技术说明：**

1、本项目使用User32.dll, Kernel32.dll所涵盖的win32 API。

2、按MSDN的API说明，win32 API支持WIN2000及以上版本，建议Win 7+。

3、Windows 7及以上可以用Python3.5+，Windows xp sp3只能用Python3.4及以下。

4、登录后的客户端最小化不影响puppet调用win32 API在后台进行交易。

**鸣谢：**

* https://github.com/hardywu/ 写了个rqalpha的接入模板PR，多谢支持！

///////////////////////做事有底线///////////////////////////////////////
