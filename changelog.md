Changelog

Puppet 1.0.1

Release date:

修复
    query('new')  # 海通证券查新股新债信息


Puppet 1.0.0

Release date: 2020-06-06

新增
    runner.run()提供一个HTTP API
    将相关的常量放入Ths类，并修改相关引用
    query()新增默认参数summary

删改
    query()的 category 参数可以为：'summary', 'position', 'order', 'deal', 'undone', 'historical_deal'
    'delivery_order', 'new', 'bingo'其中之一
    Client 改名为 Account
    code 改名为 symbol
    qty 改名为 quantity
    remainder 改名为 leftover
    num 改名为 id
    order_num 改名为 order_id
    raffle 改名为 purchase_new
    to_dict 默认值改为 True
    cancel_order 合并到 cancel
    bind 、login的返回值由 self 改为 dict
    answer、 purchase_new 的返回值由 tuple 类型改为 dict
    将全局常量MSG修改为puppet_util.Msg类，并修改相关引用
    修改初始化的设置项
    删除 ATTRS常量
    缩减 INIT常量
    删改self.account相关代码
    删除 position, deals, entrustment, cancelable, historical_deals 属性
    删除 assets, balance, free_bal, market_value, client 属性
    删除 summary()
