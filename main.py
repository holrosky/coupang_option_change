import json
import os
import sys
import time

import datetime
from chrome_manager import ChromeManager
from excel_manager import ExcelManager


class CoupangOptionChange:
    def __init__(self):
        self.chrome_manager = ChromeManager()

        self.excel_manager = ExcelManager()

        self.sheet_name = None
        self.select_sheet()

        self.seller_id = None
        self.seller_pwd = None

        self.select_account()

    def select_sheet(self):
        sheet_name_list = self.excel_manager.get_sheets_name()
        sheet_name_list.remove('설정')

        print("테스팅 대상 시트 이름을 선택해 주세요.")

        for idx, name in enumerate(sheet_name_list):
            print('[' + str(idx + 1) + '. ' + name + ']', end=' ')

        print('')

        while True:
            selected_num = input()

            if selected_num.isdigit() and len(sheet_name_list) >= int(selected_num) > 0:
                break

            else:
                print('올바른 번호를 입력해 주세요!')

        self.sheet_name = sheet_name_list[int(selected_num) - 1]
        self.excel_manager.set_sheet(self.sheet_name)

    def select_account(self):
        with open("config.json", "r", encoding='utf-8') as st_json:
            json_data = json.load(st_json)

        acc_list = []

        for each_acc in json_data['account']:
            temp_acc_list = []
            for v in each_acc.values():
                if v:
                    temp_acc_list.append(v)

            if temp_acc_list:
                acc_list.append(temp_acc_list)

        print("판매자 계정을 선택해 주세요.")

        for idx, each_acc in enumerate(acc_list):
            print('[' + str(idx + 1) + '. ' + each_acc[0] + ']', end=' ')

        print('')

        while True:
            selected_num = input()

            if selected_num.isdigit() and len(acc_list) >= int(selected_num) > 0:
                break

            else:
                print('올바른 번호를 입력해 주세요!')

        self.seller_id = acc_list[int(selected_num) - 1][0]
        self.seller_pwd = acc_list[int(selected_num) - 1][1]

    def run(self):
        while True:
            change_data = self.excel_manager.get_data()

            try:
                self.chrome_manager.open_chrome()
                self.chrome_manager.log_in(self.seller_id, self.seller_pwd)

                idx = 3

                for url, option_text, done_time in zip(change_data['url'], change_data['option_text'], change_data['done_time']):
                    if url and option_text and not done_time:
                        product_code = url.split('/')[-1]

                        self.chrome_manager.change_option(product_code, option_text)

                        now = datetime.datetime.now()
                        self.excel_manager.set_row_data(idx, now.strftime('%Y-%m-%d %H:%M'))

                    idx += 1

                self.chrome_manager.close()
                print('작업 완료! 프로그램을 종료해 주세요...')
                while True:
                    time.sleep(100)

            except Exception as e:
                exc_type, exc_obj, exc_tb = sys.exc_info()
                fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                print(e, fname, exc_tb.tb_lineno)
                self.chrome_manager.close()


if __name__ == "__main__":
    coupang_option_change = CoupangOptionChange()
    coupang_option_change.run()