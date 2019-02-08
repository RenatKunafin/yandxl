from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime
from datetime import timedelta


class Excel:
    def __init__(self, cfg, data):
        self.data = data
        self.query = data['query']
        self.name = cfg.get('excel', 'WB_NAME')
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
        wb.save(self.name)

    def write_to_wb(self):
        # Проверить есть ли такой воркшит
        # Если его нет, то создать и добавить в него строку с данными
        # затем завести для него строку на титульном воркшите
        # Если он есть, то добавить в него данные
        wb = load_workbook(self.name)
        ws_dashboard = wb[self.dashboard_ws_name]
        date = self.get_row_date()
        for d in self.data['data']:
            ws_name = self.create_ws_name(d['dimensions'])
            ws = wb[ws_name]
            if ws is None:
                ws = wb.create_sheet(ws_name)
                ws.append(self.titles)
                ws_dashboard.append([ws_name, d['metrics'][0], d['metrics'][1]])
            self.fill_row(ws, d, date)
        wb.save(self.name)
