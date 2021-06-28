import os
import getopt
import sys
import json
import argparse
from yametrics import Yametrics
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
    parser = argparse.ArgumentParser(description='Yandex metrics to Excel')
    parser.add_argument("--init", action='store_true', help="Init new report file")
    parser.add_argument("--add", action='store_true', help="Add data to existing report file")
    parser.add_argument("--date1", help="Start date")
    parser.add_argument("--date2", help="End date")
    args = parser.parse_args()
    print('!>', args)

    base_path = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(base_path, "params.ini")

    if os.path.exists(config_path):
        cfg = ConfigParser()
        cfg.read(config_path)
    else:
        print("Add params.ini")
        sys.exit(1)

    ym = Yametrics(cfg, args.date1, args.date2)
    data = ym.request_metrics()
    excel = Excel2(cfg, data, args.date1)

    if args.init is True:
        excel.init_wb()
        send(cfg)
        sys.exit()
    elif args.add is True:
        excel.write_to_wb()
        send(cfg)
        sys.exit()
    else:
        print('yandxl.py -init or -add')
        exit()

if __name__ == "__main__":
    main(sys.argv[1:])
