import requests
import pprint
import os
import logging

logging.basicConfig(level=logging.DEBUG)


class Yametrics:
    def __init__(self):
        self.url = os.environ.get('API_ROOT_URL')
        self.url_logs = os.environ.get('API_ROOT_URL2')
        self.token = os.environ.get('YANDEX_TOKEN')
        self.headers = {'Authorization': 'OAuth  ' + self.token}
        self.counter_id = os.environ.get('COUNTER')
        self.dimensions = [
            'ym:s:paramsLevel1',
            'ym:s:paramsLevel2',
            'ym:s:paramsLevel3',
            'ym:s:paramsLevel4',
            'ym:s:paramsLevel5',
            'ym:s:paramsLevel6'
        ]
        self.metrics = [
            'ym:s:visits',
            'ym:s:users',
            'ym:s:sumParams',
            'ym:s:paramsNumber',
            'ym:s:avgParams',
            'ym:s:bounceRate',
            'ym:s:pageDepth',
            'ym:s:avgVisitDurationSeconds'
        ]

    def request_metrics(self):
        params = {
            "direct_client_logins": os.environ.get('LOGIN'),
            "ids": os.environ.get('COUNTER'),
            "dimensions": ",".join(self.dimensions),
            "metrics": ",".join(self.metrics)
        }
        try:
            response = requests.get(url=self.url, headers=self.headers, params=params)
            json = response.json()
            if response.status_code == 200:
                # pprint.pprint(json)
                return json
            else:
                pprint.pprint(json)
        except Exception as e:
            print(e)
