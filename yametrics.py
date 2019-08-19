import requests
import pprint
import logging

logging.basicConfig(level=logging.DEBUG)


class Yametrics:
    def __init__(self, cfg):
        self.cfg = cfg
        self.url = cfg.get('yam', 'API_ROOT_URL')
        self.token = cfg.get('yam', 'YANDEX_TOKEN')
        self.headers = {'Authorization': 'OAuth  ' + self.token}
        self.counter_id = cfg.get('yam', 'COUNTER')
        self.period = cfg.get('yam', 'PERIOD')
        self.filters = cfg.get('yam', 'FILTERS')

        self.dimensions = cfg.get('yam', 'DIMENSIONS').split(',')
        self.metrics = cfg.get('yam', 'METRICS').split(',')

    def request_metrics(self):
        params = {
            "ids": self.counter_id,
            "dimensions": self.dimensions,
            "metrics": self.metrics,
            "date1": self.period,
            "filters": self.filters
        }
        try:
            response = requests.get(url=self.url, headers=self.headers, params=params)
            json = response.json()
            if response.status_code == 200:
                return json
            else:
                pprint.pprint(json)
        except Exception as e:
            print(e)
