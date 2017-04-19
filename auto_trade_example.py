from engine import run

basConfig = {
    "strategy_file": "./buy_and_hold.py",
    "start_date": "2016-06-01",
    "end_date": "2016-12-01",
    "stock_starting_cash": 100000,
    "benchmark": "000300.XSHG",
}

run(baseConfig)
