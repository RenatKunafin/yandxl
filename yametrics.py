import requests
import pprint
import logging
import json

logging.basicConfig(level=logging.DEBUG)


class Yametrics:
    def __init__(self, cfg, startDate, endDate):
        self.cfg = cfg
        self.url = cfg.get('yam', 'API_ROOT_URL')
        self.token = cfg.get('yam', 'YANDEX_TOKEN')
        self.headers = {'Authorization': 'OAuth  ' + self.token}
        self.counter_id = cfg.get('yam', 'COUNTER')
        self.date1 = startDate or cfg.get('yam', 'DATE1')
        self.date2 = endDate or cfg.get('yam', 'DATE2')
        self.filters = cfg.get('yam', 'FILTERS')
        self.accuracy = cfg.get('yam', 'ACCURACY')

        self.dimensions = cfg.get('yam', 'DIMENSIONS').split(',')
        self.metrics = cfg.get('yam', 'METRICS2').split(',')

    def request_metrics(self):
        print('DATES', self.date1, self.date2)
        params = {
            "ids": self.counter_id,
            "accuracy": self.accuracy,
            "dimensions": self.dimensions,
            "metrics": self.metrics,
            "date1": self.date1,
            "date2": self.date2,
            "filters": self.filters,
            "limit": 100000
        }
        try:
            response = requests.get(url=self.url, headers=self.headers, params=params)
            jsonResp = response.json()
            f = open("response.json", "w")
            json.dump(jsonResp, f)
            f.close()
            if response.status_code == 200:
                return jsonResp
            else:
                pprint.pprint(jsonResp)
        except Exception as e:
            print(e)
