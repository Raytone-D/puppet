# -*- coding: utf-8 -*-

from setuptools import setup

from puppet import __version__ as VERSION

REQUIRED = ['baidu-aip', 'keyboard']
REQUIRES_PYTHON = '>=3.4.0'

setup(
    name='puppet',
    version=VERSION,
    description=("一个用来交易A股的应用编程接口"),
    license='MIT',
    author='Raytone-D',
    author_email='raytone@qq.com',
    url='https://github.com/Raytone-D/puppet',
    keywords="stock TraderApi Quant",
    python_requires=REQUIRES_PYTHON,
    install_requires=REQUIRED,
    packages=['puppet'],
    #test_suite='tests',
    classifiers=[
        'Development Status :: 5 - Production/Stable',
        'Environment :: Win32 (MS Windows)',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
    ],
)
