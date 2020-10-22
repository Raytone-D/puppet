'''Author: Raytone-D
    Date: 2020-10-21
'''

from . import Account
from .util import find_all


class Manager:
    '''管理多个木偶实例的类'''
    def __init__(self, accno: int =None, to_dict=True, keyboard=False):
        self.acc_list = [Account(title=root, to_dict=to_dict, keyboard=keyboard) for root in find_all()]
        self.accno_list = [int(acc.id) for acc in self.acc_list]
        self.take(accno or self.accno_list[0])

    def take(self, accno: int):
        '''指定交易账号'''
        if accno in self.accno_list:
            self.acc = self.acc_list[self.accno_list.index(accno)]
            return self.acc

    def __getattr__(self, name: str):
        return getattr(self.acc, name)
