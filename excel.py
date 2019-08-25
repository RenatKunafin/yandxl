from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime


class Excel:
    def __init__(self, cfg, data):
        self.data = data
        self.query = data['query']
        self.name = cfg.get('excel', 'WB_NAME')
        self.existing_wb = cfg.get('smtp', 'PATH') + cfg.get('excel', 'WB_NAME')
        self.dashboard_ws_name = cfg.get('excel', 'DASHBOARD_WS_NAME')
        self.titles_dashboard = cfg.get('excel', 'ROW_TITLES_DASHBOARD').split(',')
        self.titles = cfg.get('excel', 'ROW_TITLES').split(',')

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
        print('REQUESTED DATES', self.query['date1'], self.query['date2'])
        date1 = datetime.strptime(self.query['date1'], '%Y-%m-%d')
        date2 = datetime.strptime(self.query['date2'], '%Y-%m-%d')
        delta = date2 - date1
        if delta.days <= 1:
            date = self.query['date1']
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
        ws_dashboard = wb.active
        ws_dashboard.title = self.dashboard_ws_name
        ws_dashboard.append(self.titles_dashboard)
        date = self.get_row_date()
        for d in self.data['data']:
            ws_name = self.create_ws_name(d['dimensions'])
            ws = wb.create_sheet(ws_name)
            ws.append(self.titles)
            self.fill_row(ws, d, date)
            ws_dashboard.append([ws_name, d['metrics'][0], d['metrics'][1]])
        wb.save(self.existing_wb)

    def write_to_wb(self):
        # Подгрузить файл, если его нет, то создать
        # Проверить есть ли воркшит с историческими данными для данной метрики
        # Если его нет, то создать и добавить в него строку с данными
        # Затем завести для него строку на титульном воркшите
        try:
            wb = load_workbook(self.existing_wb)
            ws_dashboard = wb[self.dashboard_ws_name]
            date = self.get_row_date()
            for d in self.data['data']:
                ws_name = self.create_ws_name(d['dimensions'])
                try:
                    ws = wb[ws_name]
                except KeyError:
                    ws = wb.create_sheet(ws_name)
                    ws.append(self.titles)
                    ws_dashboard.append([ws_name, d['metrics'][0], d['metrics'][1]])
                self.fill_row(ws, d, date)
            wb.save(self.name)
        except FileNotFoundError:
            self.init_wb()
