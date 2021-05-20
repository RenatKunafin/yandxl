import os
import getopt
import sys
import json
from yametrics import Yametrics
# from excel import Excel
from excel2 import Excel2
from configparser import ConfigParser
from sendmail import send_mail


def send(cfg):
    from_address = cfg.get('smtp', 'FROM')
    to_address = cfg.get('smtp', 'TO')
    subject = cfg.get('smtp', 'SUBJECT')
    text = cfg.get('smtp', 'TEXT')
    file = cfg.get('smtp', 'PATH') + cfg.get('excel', 'WB_NAME')
    password = cfg.get('smtp', 'PASS')

    send_mail(from_address, to_address, subject, text, password, [file])


def main(argv):
    try:
        opts, args = getopt.getopt(argv, "hia", ["init", "add"])
    except getopt.GetoptError as e:
        print(e)
        print('yandxl.py -i <init> -a <add>')
        sys.exit(1)

    if len(opts) == 0 and len(args) == 0:
        print('yandxl.py -i <init> -a <add>')
        sys.exit(1)
    for opt, arg in opts:
        if opt == '-h':
            print('yandxl.py -i <init> -a <add>')
            sys.exit()

    base_path = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(base_path, "params.ini")

    if os.path.exists(config_path):
        cfg = ConfigParser()
        cfg.read(config_path)
    else:
        print("Add params.ini")
        sys.exit(1)

    ym = Yametrics(cfg)
    data = ym.request_metrics()
    excel = Excel2(cfg, data)

    for opt, arg in opts:

        if opt in ('-i', '--init'):
            excel.init_wb()
            # send(cfg)
            sys.exit()
        # elif opt in ('-a', '--add'):
        #     excel.write_to_wb()
        #     # send(cfg)
        #     sys.exit()
        else:
            print('yandxl.py -i <init> -a <add>')
            exit()


if __name__ == "__main__":
    main(sys.argv[1:])
