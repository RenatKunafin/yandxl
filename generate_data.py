import os
import sys
import argparse
import requests
import pprint
from datetime import datetime
from datetime import timedelta
from configparser import ConfigParser
from excel import Excel


def main(argv):
    parser = argparse.ArgumentParser(description='Reports generaor')
    parser.add_argument("--date1", help="Start date")
    parser.add_argument("--date2", help="End date")
    args = parser.parse_args()
    print('!>', args)

    base_path = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(base_path, "params.ini")

    cfg = ConfigParser()
    cfg.read(config_path)

    url = cfg.get('yam', 'API_ROOT_URL')
    token = cfg.get('yam', 'YANDEX_TOKEN')
    headers = {'Authorization': 'OAuth  ' + token}
    counter_id = cfg.get('yam', 'COUNTER')
    filters = cfg.get('yam', 'FILTERS')
    accuracy = cfg.get('yam', 'ACCURACY')

    dimensions = cfg.get('yam', 'DIMENSIONS').split(',')
    metrics = cfg.get('yam', 'METRICS').split(',')
    params = {
            "ids": counter_id,
            "accuracy": accuracy,
            "dimensions": dimensions,
            "metrics": metrics,
            # "date1": date1,
            # "date2": date2,
            "filters": filters,
            "limit": 100000
        }

    current_date = datetime.strptime(args.date1, '%Y-%m-%d')
    end_date = datetime.strptime(args.date2, '%Y-%m-%d')
    date_step = timedelta(days=1)

    while current_date != end_date:
        print('REQUESTING: ', current_date.strftime('%d-%m-%Y'))
        try:
            current_date_string = current_date.strftime('%Y-%m-%d')
            params['date1'] = current_date_string
            params['date2'] = current_date_string
            response = requests.get(url=url, headers=headers, params=params)
            jsonResp = response.json()
            if response.status_code == 200:
                excel = Excel(cfg, jsonResp, current_date_string)
                excel.write_to_wb(config=cfg, mode='generation')
                current_date = current_date + date_step
            else:
                pprint.pprint(jsonResp)
                print('BROKE AT: ', current_date.strftime('%d-%m-%Y'))
        except Exception as e:
            print(e)



if __name__ == "__main__":
    main(sys.argv[1:])