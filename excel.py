from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime
from hashlib import md5


class Excel:
    def __init__(self, cfg, data):
        self.data = data
        self.query = data['query']
        self.path_to_wb = cfg.get('smtp', 'PATH') + cfg.get('excel', 'WB_NAME')
        self.dashboard_ws_name = cfg.get('excel', 'DASHBOARD_WS_NAME')
        self.titles_dashboard = cfg.get('excel', 'ROW_TITLES_DASHBOARD').split(',')
        self.titles = cfg.get('excel', 'ROW_TITLES').split(',')
        self.max_ws_name_length = int(cfg.get('excel', 'MAX_WS_NAME_LENGTH'))

    @staticmethod
    def fill_row(ws, data, date):
        values = list()
        values.append(date)
        values.extend(data['metrics'])
        ws.append(values)

    def create_ws_name(self, dimensions):
        name = ''
        for i, d in enumerate(dimensions):
            if i >= 2 and d['name'] is not None and d['name'] is not 'null':
                name = name + d['name'] + '.'
        a = {
            'full': name[:-1],
            'short': str(md5(name[:-1].encode('UTF-8')).hexdigest()[:-1])
        }
        return a

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
            ws = wb.create_sheet(ws_name['short'])
            ws.append(self.titles)
            self.fill_row(ws, d, date)
            ws['A1'].hyperlink = f'#{self.dashboard_ws_name}!A1'
            ws['A1'].style = "Hyperlink"
            ws_dashboard.append([ws_name['full'], d['metrics'][0], d['metrics'][1]])

        self.update_dashboard(wb)
        # wb.save(self.path_to_wb)

    def write_to_wb(self):
        # Подгрузить файл, если его нет, то создать
        # Проверить есть ли воркшит с историческими данными для данной метрики
        # Если его нет, то создать и добавить в него строку с данными
        # Затем завести для него строку на титульном воркшите
        try:
            wb = load_workbook(self.path_to_wb)
            ws_dashboard = wb[self.dashboard_ws_name]
            date = self.get_row_date()
            for d in self.data['data']:
                ws_name = self.create_ws_name(d['dimensions'])
                try:
                    ws = wb[ws_name['short']]
                except KeyError:
                    ws = wb.create_sheet(ws_name['short'])
                    ws.append(self.titles)
                    ws_dashboard.append([ws_name['full'], d['metrics'][0], d['metrics'][1]])
                    ws['A1'].hyperlink = f'#{self.dashboard_ws_name}!A1'
                    ws['A1'].style = "Hyperlink"
                self.fill_row(ws, d, date)
                wb.save(self.path_to_wb)
            self.update_dashboard(wb)
        except FileNotFoundError:
            self.init_wb()

    def get_last_row(self, ws):
        last = len(ws['A'])
        return [ws.cell(last, 2).value, ws.cell(last, 3).value]
    
    def update_dashboard(self, wb):
        ws = wb[self.dashboard_ws_name]
        for row in ws.iter_rows(min_row=ws.min_row, max_row=ws.max_row, min_col=1, max_col=3):
            for cell in row:
                if cell.row == 1:
                    continue
                elif cell.column == 1:
                    link = str(md5(cell.value.encode('UTF-8')).hexdigest()[:-1])
                    cell.hyperlink = f'#{link}!A1'
                    cell.style = "Hyperlink"
                elif cell.column == 2:
                    val = self.get_last_row(wb[str(md5(cell.offset(row=0, column=-1).value.encode('UTF-8')).hexdigest()[:-1])])
                    cell.value = val[0]
                elif cell.column == 3:
                    val = self.get_last_row(wb[str(md5(cell.offset(row=0, column=-2).value.encode('UTF-8')).hexdigest()[:-1])])
                    cell.value = val[1]
        wb.save(self.path_to_wb)