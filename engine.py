from rqalpha import run

config = {
    "base": {
        "strategy_file": "./buy_and_hold.py",
        "start_date": "2016-06-01",
        "end_date": "2016-12-01",
        "stock_starting_cash": 100000,
        "benchmark": "000300.XSHG",
    },
    "extra": {
        "log_level": "verbose",
    },
    "mod": {
      "hello_world": {
        "lib": "hello_world",
        "enabled": True,
        "priority": 100,
      }
    }
}

run(config)
