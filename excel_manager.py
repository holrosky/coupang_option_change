import json
import time

import gspread


class ExcelManager():
    def __init__(self):
        scope = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive',
        ]

        self.gc = gspread.service_account(filename='excel_credential.json', scopes=scope)

        with open("excel_credential.json", "r", encoding="UTF8") as st_json:
            json_data = json.load(st_json)

        self.doc = self.gc.open_by_url(json_data['excel_url'])
        self.sheet = None

    def get_data(self):
        data_dict = {'url': self.sheet.col_values(5)[2:], 'option_text': self.sheet.col_values(16)[2:],
                     'done_time': self.sheet.col_values(17)[2:]}

        max_len = max(len(data_dict['url']), len(data_dict['option_text']), len(data_dict['done_time']))

        while len(data_dict['url']) < max_len:
            data_dict['url'].append('')

        while len(data_dict['option_text']) < max_len:
            data_dict['option_text'].append('')

        while len(data_dict['done_time']) < max_len:
            data_dict['done_time'].append('')

        return data_dict

    def set_row_data(self, row_index, change_time):
        while True:
            try:
                change_date = self.sheet.acell('Q' + str(row_index))
                change_date.value = change_time

                self.sheet.update_cells([change_date])
                break
            except Exception as e:
                print('구글 API 요청 시간이 너무 짧습니다. 1분 후 다시 시도합니다.')
                time.sleep(60)

    def set_sheet(self, sheet_name):
        self.sheet = self.doc.worksheet(sheet_name)

    def get_sheets_name(self):
        return [s.title for s in self.doc.worksheets()]
