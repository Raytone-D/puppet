import rqalpha

config = {
    "extra": {
        "log_level": "verbose",
    },
    "mod": {
      "live_trade": {
        "lib": "./mod",
        "enabled": True,
        "priority": 100,
      }
    }
}

def run(baseConf):
  config["base"] = baseConf
  return rqalpha.run(config)
