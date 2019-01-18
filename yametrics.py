import requests
import pprint
import os
import logging

logging.basicConfig(level=logging.DEBUG)


class Yametrics:
    def __init__(self):
        self.url = os.environ.get('API_ROOT_URL')
        self.token = os.environ.get('YANDEX_TOKEN')
        self.headers = {'Authorization': 'OAuth  ' + self.token}
        self.metrics = ['ym:s:visits', 'ym:s:pageviews', 'ym:s:users']

    def request_metrics(self):
        params = {
            "direct_client_logins": os.environ.get('LOGIN'),
            "ids": os.environ.get('COUNTER'),
            "metrics": ",".join(self.metrics)
        }
        try:
            response = requests.get(url=self.url, headers=self.headers, params=params)
            json = response.json()
            print(response)
            if response.status_code == 200:
                pprint.pprint(json)
            else:
                print('Error')
        except Exception as e:
            print(e)
