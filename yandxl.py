from yametrics import Yametrics
from excel import Excel
import os
import sys
from configparser import ConfigParser
from sendmail import send_mail

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
e.write_to_wb()

server = cfg.get('smtp', 'SERVER')
port = cfg.get('smtp', 'PORT')
from_address = cfg.get('smtp', 'FROM')
to_address = cfg.get('smtp', 'TO').split(',')
subject = cfg.get('smtp', 'SUBJECT')
text = cfg.get('smtp', 'TEXT')
file = cfg.get('smtp', 'PATH') + cfg.get('excel', 'WB_NAME')
password = cfg.get('smtp', 'PASS')

send_mail(from_address, to_address, subject, text, password, [file])
