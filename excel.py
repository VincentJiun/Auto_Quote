from openpyxl import Workbook, load_workbook
import os

class Excel():
    def __init__(self, path):
        self.path = f'./data/{path}'
        if os.path.exists(self.path):
            self.wb = load_workbook(self.path)
        else:
            self.wb = Workbook()
            self.wb.save(self.path)  # 只在創建新工作簿時保存

class ExcelCMS(Excel):
    def __init__(self, path):
        super().__init__(path)

    def create(self, *args):
        self.ws = self.wb.active
        self.ws.append(args)
        self.wb.save(self.path)  # 保存檔案
        # self.wb.close()  # 不要在這裡關閉，保留後續使用

    def read_all_datas(self):
        data = []
        self.ws = self.wb.active
        for row in self.ws.iter_rows(values_only=True):
            # 將每一行的資料轉換為列表，並添加到 data 中
            data.append(list(row))

        # 關閉工作簿
        self.wb.close()

        return data



