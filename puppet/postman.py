'''向木偶投递交易指令
Release date: 2020-09-10
'''
__author__ = "睿瞳深邃(https://github.com/Raytone-D)"
__license__ = 'MIT'

try:
    from httpx import Client as Session
except Exception:
    from requests import Session


url = 'http://127.0.0.1:10086/puppet'
_s = Session()


def set_token(acc_id, sault):
    '''acc_id 是 puppet.Account 实例的 id, sault 是 puppet.postman.sault'''
    token = acc_id
    globals().append(s=Session())


def buy(symbol: str, price: float, quantity: int =100, action: str ='buy') -> dict:
    '''Limit Order.
    Return:
        Response.status_code, Response.json()'''
    return _s.post(url, json=locals()).json()


def sell(symbol: str, price: float, quantity: int =100, action: str ='sell') -> dict:
    '''Limit Order.'''
    return _s.post(url, json=locals()).json()


def cancel_all(action: str ='cancel_all') -> dict:
    '''Args:
        action: 'cancel_all' or 'margin_cancel_all'
    '''
    return _s.post(url, json=locals()).json()


def query(category: str ='summary', action: str ='query') -> dict:
    return _s.post(url, json=locals()).json()


def bind(title: str='', action: str ='bind') -> dict:
    return _s.post(url, json=locals()).json()


def login(account_no: str, password: str, comm_pwd: str, client_path: str, \
    token: str, action: str ='login') -> dict:
    return _s.post(url, json=locals()).json()
