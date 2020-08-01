'''
Author: 睿瞳深邃(https://github.com/Raytone-D)
License: MIT
Release date: 2020-06-06
Version: 0.3
'''

from .client import Account, __version__


def run(host='127.0.0.1', port=10086):
    '''Puppet's HTTP Service'''
    from bottle import post, run, request, response

    @post('/puppet')
    def serve():
        '''Puppet Web Trading Interface'''
        task = request.json
        if task:
            try:
                return getattr(acc, task.pop('action'))(**task)
            except Exception as e:
                response.bind(status=502)
                return {'puppet': str(e)}
        return {'puppet': '仅支持json格式'}

    print('Puppet version:', __version__)
    acc = Account()
    run(host=host, port=port)


if __name__ == "__main__":
    run()
