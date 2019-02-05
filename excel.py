from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime
from datetime import timedelta
import os


class Excel:
    def __init__(self, data):
        self.data = data
        self.query = data['query']
        self.name = os.environ.get('WB_NAME')
        # TODO вынести в .envrc
        self.titles = [
            'Дата',
            'Визиты',
            'Пользователи',
            'Сумма параметров визитов',
            'Количество параметров визитов',
            'Среднее параметров визитов',
            'Отказы',
            'Глубина просмотра',
            'Время на сайте'
        ]

    @staticmethod
    def create_ws_name(dimensions):
        name = ''
        for i, d in enumerate(dimensions):
            if i > 2 and d['name'] is not None:
                name = name + d['name'] + '.'
        return name[:-1]

    @staticmethod
    def fill_row(ws, data, date):
        values = list()
        values.append(date)
        values.extend(data['metrics'])
        ws.append(values)

    def get_row_date(self):
        date = ''
        if self.query['date1'] is 'today':
            date = datetime.now().strftime("%Y-%m-%d")
        elif self.query['date1'] is 'yesterday':
            date = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        else:
            try:
                date1 = datetime.strptime(self.query['date1'], '%Y-%m-%d').strftime("%Y-%m-%d")
                date2 = datetime.strptime(self.query['date2'], '%Y-%m-%d').strftime("%Y-%m-%d")
                date = str(date1) + ' - ' + str(date2)
            except ValueError as e:
                print(e)
        return str(date)

    def init_wb(self):
        wb = Workbook()
        date = self.get_row_date()
        for d in self.data['data']:
            ws = wb.create_sheet(Excel.create_ws_name(d['dimensions']))
            ws.append(self.titles)
            self.fill_row(ws, d, date)
        wb.save(self.name)

    def write_to_wb(self):
        wb = load_workbook(self.name)
        date = self.get_row_date()
        for d in self.data['data']:
            ws = wb[self.create_ws_name(d['dimensions'])]
            self.fill_row(ws, d, date)
        wb.save(self.name)
