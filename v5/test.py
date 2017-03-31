""" 单客户端测试 """

# coding: utf-8

from puppet_v5 import Puppet

trader = Puppet()

print(trader.position)
print(trader.balance)
print(trader.cancelable)
print(trader.deals)
print(trader.new)
print(trader.bingo)
#trader.raffle()
#trader.sell(300111, 4.99, 100)    # 已经无需字符串。
