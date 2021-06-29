from openpyxl import Workbook, load_workbook
from openpyxl.cell import Cell
from openpyxl.styles import Font, Fill, PatternFill
from datetime import datetime
from datetime import timedelta
from hashlib import md5
import time
import json
import re

from openpyxl.workbook.defined_name import ROW_RANGE_RE


class Excel:
    def __init__(self, cfg, data, startDate):
        self.data = data
        # self.path_to_wb = cfg.get('smtp', 'PATH') + cfg.get('excel', 'WB_NAME') + '_' + str(time.time()) + '.xlsx'
        self.path_to_wb = cfg.get('smtp', 'PATH') + cfg.get('excel', 'WB_NAME')
        self.dashboard_ws_name = cfg.get('excel', 'DASHBOARD_WS_NAME')
        self.odbo_ws_name = cfg.get('excel', 'ODBO_FUNNEL_SHEET_NAME')
        self.odbo_titles = cfg.get('excel', 'ODBO_FUNNEL_TITLES').split(',')
        self.odbo_steps = cfg.get('excel', 'ODBO_FUNNEL_STEPS').split(',')
        self.titles_dashboard = cfg.get('excel', 'ROW_TITLES_DASHBOARD').split(',')
        self.date1 = startDate or cfg.get('yam', 'DATE1')
        self.title_font = Font(size=14, color='FF000000')

    def _build_dashboard_row(self, dimensions, metrics):
        entry = []
        for d in dimensions:
            entry.append(d['name'])
        return entry + metrics

    def _build_dashboard(self, wb):
        if not self.dashboard_ws_name in wb:
            ws_dashboard = wb.active
            ws_dashboard.title = self.dashboard_ws_name
            ws_dashboard.append(self.titles_dashboard)
            ws_dashboard.auto_filter.ref = ws_dashboard.dimensions
            for row in ws_dashboard.iter_rows(min_row=1, max_row=1, min_col=1, max_col=ws_dashboard.max_column):
                for cell in row:
                    cell.font = self.title_font
            ws_dashboard.column_dimensions['A'].width = 15
            ws_dashboard.column_dimensions['B'].width = 20
            ws_dashboard.column_dimensions['C'].width = 30
            ws_dashboard.column_dimensions['D'].width = 30
            ws_dashboard.column_dimensions['E'].width = 40
            ws_dashboard.column_dimensions['F'].width = 40
            ws_dashboard.column_dimensions['G'].width = 25
            ws_dashboard.column_dimensions['H'].width = 20
            ws_dashboard.column_dimensions['I'].width = 20
        ws_dashboard = wb[self.dashboard_ws_name]
        ws_dashboard.delete_rows(2, ws_dashboard.max_row)
        entries = []
        for d in self.data['data']:
            entries.append(self._build_dashboard_row(d['dimensions'], d['metrics']))
        for row in entries:
            ws_dashboard.append(row)

    def _dataToList(self, dimensions, metrics):
        return [dimensions[4]['name'], dimensions[0]['name'], dimensions[3]['name'], metrics[0]]

    def _sort(self, items):
        res = list()
        for pos in self.odbo_steps:
            for item in items:
                if pos == item[0]:
                    res.append(item)
        return res

    def _build_odbo_funnel(self):
        tree = {}
        for d in self.data['data']:
            if (d['dimensions'][2]['name'] == 'Открытие брокерского счета'):
                step = d['dimensions'][4]['name']
                channel = d['dimensions'][0]['name']
                source = d['dimensions'][3]['name']
                action = d['dimensions'][5]['name']
                value = d['metrics'][0]
                key = step+'-'+channel+'-'+source
                if action == 'Открытие' or (step == 'Экран запрета' and d['dimensions'][6]['name'] == 'Открытие'):
                    tree[key] = [step, channel, source, value]
                elif action == 'Один телефон' or action == 'Несколько телефонов':
                    try:
                        tree[key][3] = tree[key][3] + value
                    except KeyError:
                        tree[key] = [step, channel, source, value]
        self._save_to('new_data.json', tree)
        temp = list(tree.values())
        res = list()
        for pos in self.odbo_steps:
            for item in temp:
                if pos == item[0]:
                    res.append(item)
        return res

    def _get_date(self):
        p = re.compile('(\d+)[A-z]+')
        res = p.findall(self.date1)
        if len(res) != 0:
            d = datetime.today() - timedelta(days=int(res[0]))
            return d.strftime('%d-%m-%Y')
        elif self.date1 is not None:
            return datetime.strptime(self.date1, '%Y-%m-%d').strftime('%d-%m-%Y')
        else:
            raise ValueError('Unknown date1 value', self.date1)

    def _save_to(self, file_name, data):
        f = open(file_name, "w")
        f.write(json.dumps(data))
        f.close()

    def write_to_wb(self):
        print('WRITE')
        try:
            wb = load_workbook(self.path_to_wb)
            self._build_dashboard(wb)
            ws = wb[self.odbo_ws_name]
            date = self._get_date()
            column = ws.max_column
            if column == 18:
                ws.delete_cols(4)
            tree = self._build_odbo_funnel()
            old_data = {}
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column, values_only=True):
                old_data[row[0]+'-'+row[1]+'-'+row[2]] = list(row)
            self._save_to('old_data.json', old_data)

            for item in tree:
                key_from_new = item[0]+'-'+item[1]+'-'+item[2]
                value = item[3]
                if key_from_new in old_data:
                    old_data[key_from_new].append(value)
                else:
                    res = [0] * (ws.max_column - 3)
                    res.append(value)
                    temp = item[0:3] + res
                    old_data[key_from_new] = temp
            result = self._sort(list(old_data.values()))
            self._save_to('result.json', result)
            ws.delete_rows(2, ws.max_row)
            for row in result:
                ws.append(row)
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=ws.max_column, max_col=ws.max_column):
                for cell in row:
                    if cell.value is None:
                        cell.value = 0
            ws.cell(row = 1, column=ws.max_column).value = date
            ws.cell(row = 1, column=ws.max_column).font = self.title_font
            ws.column_dimensions[ws.cell(row = 1, column=ws.max_column).column_letter].width = 15
            wb.save(self.path_to_wb)
            print('excel ready')
            
        except FileNotFoundError:
            print('Excel file not found')

    def init_wb(self):
        print('INIT')
        wb = Workbook()
        self._build_dashboard(wb)
        ws_odbo = wb.create_sheet(self.odbo_ws_name)
        ws_odbo.title = self.odbo_ws_name
        ws_odbo.append(self.odbo_titles)
        for row in ws_odbo.iter_rows(min_row=1, max_row=1, min_col=1, max_col=ws_odbo.max_column):
                for cell in row:
                    cell.font = self.title_font
        ws_odbo.column_dimensions['A'].width = 20
        ws_odbo.column_dimensions['B'].width = 15
        ws_odbo.column_dimensions['C'].width = 30
        ws_odbo.column_dimensions['D'].width = 15
        ws_odbo.auto_filter.ref = ws_odbo.dimensions
        tree = self._build_odbo_funnel()
        for row in tree:
            ws_odbo.append(row)
        ws_odbo.cell(row = 1, column=ws_odbo.max_column).value = self._get_date()
        ws_odbo.cell(row = 1, column=ws_odbo.max_column).font = self.title_font
        wb.save(self.path_to_wb)
        print('excel ready')
