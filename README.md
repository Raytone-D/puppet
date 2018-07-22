『扯线木偶』Puppet traderAPI
==

实现交易接口，突破交易桎梏！
--
实现了和股票交易客户端相同的【买卖撤查】功能。暂不支持【融资融券】交易功能。
--

puppet扯线木偶目前仅适用于独立交易端核新版(即THS)，需配合交易框架(hikyuu, quantaxis, rqalpha, vnpy, etc)使用。
-
墙裂推荐使用实盘易，一站式量化交易解决方案 http://www.iguuu.com/e?x=19829 。**
-

**未实现的功能：**

【验证码】登录

【市价】委托

【逆回购】

【基金盘后业务】

**已实现的登录方法：**

login()         登录客户端
exit()          关闭客户端

**已实现的交易方法：**

send_order()    双向委托，【限价】委托下单，未实现【市价】委托。
buy()           【限价】买入，未实现【市价】委托。
sell()          【限价】卖出，未实现【市价】委托。
cancel_order()  撤单
cancel_all()    全撤
cancel_buy()    撤买
cancel_sell()   撤卖
raffle()        新股申购

**已实现的查询功能：**

account         登录账号
balance         可用余额
position        持仓列表
market_value    持仓市值
cancelable      可撤委托
entrustment     当日委托
deals           当日成交
new             当日新股
bingo           中签查询，部分券商可查。


**代码示范：**

import puppet as api

p = api.Puppet()
p.login(account_no='你的账号', password='你的交易密码', comm_pwd='你的通讯密码') # 登录客户端

p.account                       # 查看当前登录的账号
p.balance                       # 查看当前账户可用余额
p.market_value                  # 当前账号的实时市值
p.buy('000001', '9.32', '100')  # 限价委托，在[9.32]这个价位买入[100股][平安银行]，注意是str类型
p.entrustment                   # 查看上述委托是否受理或成交了。
p.cancel_buy()                  # 撤销当前全部买单


**使用环境：**

1、同花顺 股票交易客户端，留意【部分最新】版本不能复制持仓或委托数据。推荐备用版本: https://pan.baidu.com/s/1radY1fI 密码: m9db

1、要求用Python3.4及以上，强烈推荐用于科学计算的Python发行版【Anaconda3】的【最新】版本。

2、Win 2k+，Windows下免安装、零配置

3、Linux需安装最新的Wine，环境设为WIN7，并安装Win平台的Anaconda3。

4、下载Anaconda3推荐用清华镜像：https://mirrors.tuna.tsinghua.edu.cn/anaconda/archive/


**试用步骤：**

1、上官网 https://github.com/Raytone-D/puppet 点击页面右上角的绿色按钮[Clone or download]，
    选择[Download ZIP]，将文件下载到本地。

2、解压文件。

3、打开命令提示符[Command Prompt]终端。切换到puppet\puppet文件夹。

4、测试：手动登录客户端，然后在终端窗口输入python puppet.py，回车。看输出没出错信息即可。


**技术说明：**

1、本项目使用User32.dll, Kernel32.dll所涵盖的win32 API。

2、按MSDN的API说明，win32 API支持WIN2000及以上版本，建议Win 7+。

3、Windows 7及以上可以用Python3.5+，Windows xp sp3只能用Python3.4及以下。

4、登录后的客户端最小化不影响puppet调用win32 API在后台进行交易。

**鸣谢：**

* https://github.com/hardywu/ 写了个rqalpha的接入模板PR，多谢支持！

///////////////////////做事有底线///////////////////////////////////////
