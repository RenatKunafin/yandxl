from yametrics import Yametrics
from excel import Excel
import os
import sys
from configparser import ConfigParser

base_path = os.path.dirname(os.path.abspath(__file__))
config_path = os.path.join(base_path, "params.ini")

if os.path.exists(config_path):
    cfg = ConfigParser()
    cfg.read(config_path)
else:
    print("Config not found! Exiting!")
    sys.exit(1)

ym = Yametrics(cfg)
data = ym.request_metrics()

e = Excel(cfg, data)
e.init_wb()
# e.write_to_wb()
